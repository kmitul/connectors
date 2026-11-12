# Custom SharePoint connector — 9.5.6

This branch is based on upstream `v9.5.6`
(`e2965155a457ca6faff613f274c801f3c778ade8`). It ports the SharePoint
customizations from `custom-sharepoint-9.3` commits `b4e92f1d` and `c9c757e2`
onto the genuine 9.5.6 connector implementation. Package versions and dependencies
remain upstream 9.5.6; this is not a retagged 9.3 runtime.

## Custom behavior retained

Only the built-in SharePoint Online `client.py` and `datasource.py` change:

- Entra OAuth2 client credentials for the remaining SharePoint REST token path.
- Microsoft Graph site-page ingestion with the existing document field names.
- Site/list metadata, hidden-list filtering, and full list-item field expansion.
- Cached M365 backing-group discovery and owners/members ACL identities.
- Graph-based resolution of drive permission `siteGroup` identities, with existing
  REST fallback.
- Batched `listItem/fields` enrichment of nondeleted drive items under
  `list_item_fields`, including with DLS or permission fetching disabled.
- Existing REST permission-error fallback behavior.

Upstream 9.5.6 ancestor-permission and expired-delta handling are retained.

## Build safely

Use a clean checkout or a tracked-files-only build context. Do not build from a
checkout containing local configuration/credentials with broad `COPY .`.
From this branch, the following sends only committed files to Docker:

```bash
git archive HEAD | docker build -t elastic-connectors:custom-9.5.6 -
docker run --rm --entrypoint /app/bin/elastic-ingest \
  elastic-connectors:custom-9.5.6 --version
```

Run with the existing connector configuration mounted read-only at
`/app/config.yml`, retaining its network and environment requirements:

```text
/app/bin/elastic-ingest --config-file /app/config.yml
```

For Python diagnostics, use
`/app/app/connectors_service/.venv/bin/python`, rather than the `/app/bin/python`
directory symlink, which can lose virtualenv package discovery.

## Focused offline regression tests

Five async tests cover M365 owners/members ACLs, Graph site-group permission
propagation, custom fields with DLS off/on, Graph page shape, and expired-delta
cursor clearing. Against a built image:

```bash
docker run --rm --network none \
  --entrypoint /app/app/connectors_service/.venv/bin/python \
  --mount type=bind,src="$PWD/app/connectors_service/tests/sources",dst=/regression-tests,readonly \
  elastic-connectors:custom-9.5.6 \
  -m unittest discover -s /regression-tests -p test_sharepoint_online_overlay.py -v
```

Validation on 2026-11-12: all five tests passed against the official 9.5.6
runtime with these exact two modules overlaid. The deployed custom image's
executable/package reports 9.5.6, and its read-only version/configuration
preflight against Elasticsearch 9.5.6 passed. The old deployed custom modules
were hash-matched to the custom-9.3 source before porting.

This is **not** a successful full upstream test suite or a completed live tenant
sync. Unchanged upstream tests/fixtures contain REST-based assumptions that differ
from this custom Graph behavior. Validate a real sync and restricted-document
access before a broad worker rollout.

## Inherited ACL limitations

The version upgrade preserves the requested existing ACL behavior; it does not
redesign the permission model:

- Any SharePoint `siteGroup` ID with a cached backing M365 group maps to that
  group's owners/members. Owners, Members, Visitors, and arbitrary custom
  SharePoint groups are not independently distinguished by this custom mapping.
- Sites without a backing M365 group can return empty site-level ACLs; Visitors
  are not independently resolved by the owners/members model.
- Some REST permission failures inherit parent permissions, potentially broader
  than unique item permissions. In other paths, a unique-permission flag already
  set before role enumeration fails can prevent the advertised fallback.
- Unique permissions, attachments, and some group fallbacks still use REST.
  Tenant permission grants must remain valid; a runtime upgrade does not grant
  Entra/Graph/SharePoint application permissions.

Correcting these limitations needs a separate permission-model change and a
controlled ACL re-sync. Existing workers do not change merely because this branch
or a new image exists; recreate each worker deliberately, without duplicate
workers for the same connector ID. Keep the previous connector image for
connector-only recovery, not for downgrading Elasticsearch data.
