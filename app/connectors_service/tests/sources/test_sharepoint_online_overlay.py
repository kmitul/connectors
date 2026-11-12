import logging
import unittest
from types import SimpleNamespace

from connectors.access_control import ACCESS_CONTROL
from connectors.sources.sharepoint.sharepoint_online.client import SharepointOnlineClient
from connectors.sources.sharepoint.sharepoint_online.datasource import (
    SharepointOnlineDataSource,
)
from connectors.sources.sharepoint.sharepoint_online.utils import DeltaLinkExpired
from connectors.sources.sharepoint.sharepoint_online.constants import (
    CURSOR_SITE_DRIVE_KEY,
)


async def _iterate(values):
    for value in values:
        yield value


class FakeClient:
    def __init__(self, owners=None, members=None, field_responses=None, permissions=None):
        self.owners = owners or []
        self.members = members or []
        self.field_responses = field_responses or []
        self.permissions = permissions or []
        self.field_batches = []
        self.permission_batches = []
        self.rest_group_fetches = 0

    async def group_owners(self, group_id):
        async for owner in _iterate(self.owners):
            yield owner

    async def group_members(self, group_id):
        async for member in _iterate(self.members):
            yield member

    async def site_group_id(self, site_name):
        return "group-from-name"

    async def drive_items_list_item_fields_batch(self, drive_id, item_ids):
        self.field_batches.append((drive_id, list(item_ids)))
        async for response in _iterate(self.field_responses):
            yield response

    async def drive_items_permissions_batch(self, drive_id, item_ids):
        self.permission_batches.append((drive_id, list(item_ids)))
        async for response in _iterate(self.permissions):
            yield response

    async def site_groups_users(self, site_web_url, site_group_id):
        self.rest_group_fetches += 1
        if False:
            yield None


class FakeGraphClient:
    def __init__(self):
        self.urls = []

    async def scroll(self, url):
        self.urls.append(url)
        if "/lists?" in url:
            yield [{"id": "pages-list-id"}]
        else:
            yield [
                {
                    "id": "page-item-id",
                    "createdDateTime": "2026-01-02T03:04:05Z",
                    "lastModifiedDateTime": "2026-02-03T04:05:06Z",
                    "webUrl": "https://contoso.sharepoint.com/sites/team/SitePages/Home.aspx",
                    "fields": {
                        "Title": "Home",
                        "CanvasContent1": "<div>canvas</div>",
                        "WikiField": "<p>wiki</p>",
                        "AuthorLookupId": "17",
                        "EditorLookupId": "19",
                        "_UIVersionString": "3.0",
                    },
                }
            ]


class SharePointOverlayTests(unittest.IsolatedAsyncioTestCase):
    def make_datasource(self, client, dls_enabled=True, fetch_item_permissions=True):
        datasource = object.__new__(SharepointOnlineDataSource)
        datasource.configuration = {
            "use_document_level_security": dls_enabled,
            "fetch_drive_item_permissions": fetch_item_permissions,
            "expand_site_group_members": True,
        }
        datasource._features = SimpleNamespace(
            document_level_security_enabled=lambda: True
        )
        datasource._logger = logging.getLogger("sharepoint-overlay-test")
        datasource._client = client
        datasource.site_group_cache = {}
        datasource._site_m365_group_cache = {}
        return datasource

    async def test_site_acl_uses_m365_owners_members_and_group_tokens(self):
        client = FakeClient(
            owners=[
                {
                    "id": "owner-id",
                    "userPrincipalName": "owner@contoso.com",
                    "mail": "owner@contoso.com",
                }
            ],
            members=[
                {
                    "id": "member-id",
                    "@odata.type": "#microsoft.graph.user",
                    "userPrincipalName": "member@contoso.com",
                    "mail": "member@contoso.com",
                },
                {
                    "id": "nested-group-id",
                    "@odata.type": "#microsoft.graph.group",
                },
            ],
        )
        datasource = self.make_datasource(client)
        site = {
            "id": "site-id",
            "webUrl": "https://contoso.sharepoint.com/sites/team",
            "name": "team",
            "group": {"id": "m365-group-id"},
        }

        acl, admins, editors = await datasource._site_access_control(site)

        self.assertIn("user:owner@contoso.com", acl)
        self.assertIn("user_id:owner-id", acl)
        self.assertIn("user:member@contoso.com", acl)
        self.assertIn("email:member@contoso.com", acl)
        self.assertIn("user_id:member-id", acl)
        self.assertIn("group:nested-group-id", acl)
        self.assertIn("group:m365-group-id", acl)
        self.assertIn("user:owner@contoso.com", admins)
        self.assertNotIn("user:member@contoso.com", admins)
        self.assertEqual(set(acl), set(editors))

    async def test_site_group_permission_propagates_via_graph_without_rest(self):
        client = FakeClient(
            owners=[{"id": "owner-id", "userPrincipalName": "owner@contoso.com"}],
            members=[{"id": "member-id", "userPrincipalName": "member@contoso.com"}],
        )
        datasource = self.make_datasource(client)
        datasource._site_m365_group_cache["site-id"] = "m365-group-id"
        item = {"id": "drive-item-id"}

        result = await datasource._with_drive_item_permissions(
            item,
            [{"grantedToV2": {"siteGroup": {"id": "2"}}}],
            "https://contoso.sharepoint.com/sites/team",
            site_id="site-id",
        )

        self.assertIn("user:member@contoso.com", result[ACCESS_CONTROL])
        self.assertIn("user:owner@contoso.com", result[ACCESS_CONTROL])
        self.assertEqual(client.rest_group_fetches, 0)

    async def test_drive_item_fields_are_enriched_with_dls_disabled_and_enabled(self):
        for dls_enabled in (False, True):
            with self.subTest(dls_enabled=dls_enabled):
                client = FakeClient(
                    field_responses=[
                        {
                            "id": "item-id",
                            "status": 200,
                            "body": {"CustomText": "indexed value"},
                        }
                    ],
                    permissions=[
                        {
                            "id": "item-id",
                            "body": {
                                "value": [
                                    {"grantedToV2": {"user": {"id": "user-id"}}}
                                ]
                            },
                        }
                    ],
                )
                datasource = self.make_datasource(client, dls_enabled=dls_enabled)
                items = [{"id": "item-id", "name": "document.docx"}]

                results = [
                    item
                    async for item in datasource._drive_items_batch_with_permissions(
                        "drive-id", items, "https://contoso.sharepoint.com/sites/team"
                    )
                ]

                self.assertEqual(
                    results[0]["list_item_fields"], {"CustomText": "indexed value"}
                )
                self.assertEqual(client.field_batches, [("drive-id", ["item-id"])])
                if dls_enabled:
                    self.assertEqual(
                        client.permission_batches, [("drive-id", ["item-id"])]
                    )
                    self.assertIn("user_id:user-id", results[0][ACCESS_CONTROL])
                else:
                    self.assertEqual(client.permission_batches, [])
                    self.assertNotIn(ACCESS_CONTROL, results[0])

    async def test_graph_site_page_shape_matches_existing_ingestion_keys(self):
        client = object.__new__(SharepointOnlineClient)
        client._graph_api_client = FakeGraphClient()
        client._logger = logging.getLogger("sharepoint-page-test")

        pages = [page async for page in client.site_pages("site-id")]

        self.assertEqual(len(pages), 1)
        page = pages[0]
        self.assertEqual(page["Id"], "page-item-id")
        self.assertEqual(page["Title"], "Home")
        self.assertEqual(
            page["webUrl"],
            "https://contoso.sharepoint.com/sites/team/SitePages/Home.aspx",
        )
        self.assertEqual(page["Created"], "2026-01-02T03:04:05Z")
        self.assertEqual(page["Modified"], "2026-02-03T04:05:06Z")
        self.assertEqual(page["AuthorId"], "17")
        self.assertEqual(page["EditorId"], "19")
        self.assertEqual(page["OData__UIVersionString"], "3.0")
        self.assertTrue(any("pages-list-id" in url for url in client._graph_api_client.urls))

    async def test_upstream_expired_delta_clears_cursor_after_recovery_failure(self):
        client = FakeClient()

        async def expired_drive_items(drive_id, url=None):
            if False:
                yield None
            raise DeltaLinkExpired("expired")

        client.drive_items = expired_drive_items
        datasource = self.make_datasource(client)
        datasource._sync_cursor = {
            CURSOR_SITE_DRIVE_KEY: {"drive-id": "expired-delta-link"}
        }

        with self.assertRaises(DeltaLinkExpired):
            async for _ in datasource._sync_drive_items("drive-id"):
                pass

        self.assertEqual(datasource._sync_cursor[CURSOR_SITE_DRIVE_KEY], {})


if __name__ == "__main__":
    unittest.main()
