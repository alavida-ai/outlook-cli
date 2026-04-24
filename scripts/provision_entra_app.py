"""Register (or update) the Entra app for the Alavida Outlook CLI.

Run once to bootstrap the shared multi-tenant app. Idempotent: if an app with the same
display name exists, updates its `signInAudience` in place instead of creating a duplicate.

Requires the running user to be a Global Admin (or Application Administrator) on the tenant.
First invocation opens a browser for interactive sign-in.

Usage:
    # First-time bootstrap of the shared multi-tenant app (recommended):
    uv run python scripts/provision_entra_app.py --tenant alavidai.onmicrosoft.com --multi-tenant

    # Paranoid-client escape hatch: dedicated single-tenant app in the client's own tenant.
    uv run python scripts/provision_entra_app.py --tenant <client-tenant>
"""

from __future__ import annotations

import argparse
import asyncio
import sys

from azure.identity import InteractiveBrowserCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.application import Application
from msgraph.generated.models.public_client_application import PublicClientApplication
from msgraph.generated.models.required_resource_access import RequiredResourceAccess
from msgraph.generated.models.resource_access import ResourceAccess

APP_DISPLAY_NAME = "alavida-outlook-cli"
GRAPH_RESOURCE_APP_ID = "00000003-0000-0000-c000-000000000000"
NATIVE_REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"

DELEGATED_SCOPES: list[tuple[str, str]] = [
    ("024d486e-b451-40bb-833d-3e66d98c5c73", "Mail.ReadWrite"),
    ("1ec239c2-d7c9-4623-a91a-a9775856bb36", "Calendars.ReadWrite"),
    ("12466101-c9b8-439a-8589-dd09ee67e8e9", "Calendars.ReadWrite.Shared"),
    ("d56682ec-c09e-4743-aaf4-1a3aac4caa21", "Contacts.ReadWrite"),
    ("7427e0e9-2fba-42fe-b0c0-848c9e6a8182", "offline_access"),
    ("e1fe6dd8-ba31-4d61-89e7-88639da4683d", "User.Read"),
]


def _build_app(multi_tenant: bool) -> Application:
    return Application(
        display_name=APP_DISPLAY_NAME,
        sign_in_audience="AzureADMultipleOrgs" if multi_tenant else "AzureADMyOrg",
        public_client=PublicClientApplication(redirect_uris=[NATIVE_REDIRECT_URI]),
        is_fallback_public_client=True,
        required_resource_access=[
            RequiredResourceAccess(
                resource_app_id=GRAPH_RESOURCE_APP_ID,
                resource_access=[
                    ResourceAccess(id=scope_id, type="Scope") for scope_id, _ in DELEGATED_SCOPES
                ],
            ),
        ],
    )


async def _find_existing(client: GraphServiceClient) -> Application | None:
    from msgraph.generated.applications.applications_request_builder import (
        ApplicationsRequestBuilder,
    )

    qp = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetQueryParameters(
        filter=f"displayName eq '{APP_DISPLAY_NAME}'",
    )
    config = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetRequestConfiguration(
        query_parameters=qp,
    )
    page = await client.applications.get(request_configuration=config)
    if page and page.value:
        return page.value[0]
    return None


async def _run(tenant_id: str, multi_tenant: bool) -> None:
    credential = InteractiveBrowserCredential(tenant_id=tenant_id)
    client = GraphServiceClient(
        credentials=credential,
        scopes=["https://graph.microsoft.com/.default"],
    )

    desired_audience = "AzureADMultipleOrgs" if multi_tenant else "AzureADMyOrg"
    existing = await _find_existing(client)

    if existing is not None:
        app = existing
        print(f"App '{APP_DISPLAY_NAME}' already exists (object id {app.id}).")
        if app.sign_in_audience != desired_audience:
            print(
                f"Updating signInAudience: {app.sign_in_audience} -> {desired_audience}"
            )
            patch = Application(sign_in_audience=desired_audience)
            await client.applications.by_application_id(app.id).patch(patch)
            app.sign_in_audience = desired_audience
            print("Updated.")
        else:
            print(f"signInAudience already {desired_audience} — nothing to change.")
    else:
        print(f"Creating app '{APP_DISPLAY_NAME}' ({desired_audience})...")
        app = await client.applications.post(_build_app(multi_tenant))
        print("Created.")

    print()
    print("=" * 60)
    print("Embed this client ID in src/outlook_cli/auth.py:")
    print(f"  DEFAULT_CLIENT_ID = \"{app.app_id}\"")
    print()
    print("Or keep it in .env (overrides the embedded default):")
    print(f"  AZURE_CLIENT_ID={app.app_id}")
    if not multi_tenant:
        print(f"  AZURE_TENANT_ID={tenant_id}   # only needed for single-tenant apps")
    print()
    print("Grant admin consent for THIS tenant at:")
    consent_url = (
        f"https://login.microsoftonline.com/{tenant_id}/adminconsent"
        f"?client_id={app.app_id}"
    )
    print(f"  {consent_url}")
    if multi_tenant:
        print()
        print("For each client tenant, send their admin this URL template:")
        print(
            f"  https://login.microsoftonline.com/<THEIR-TENANT-ID>/adminconsent"
            f"?client_id={app.app_id}"
        )
    print()
    print("Scopes requested:")
    for _, name in DELEGATED_SCOPES:
        print(f"  - {name}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tenant", required=True, help="Tenant domain or ID (e.g. alavidai.onmicrosoft.com)")
    parser.add_argument(
        "--multi-tenant",
        action="store_true",
        help="Register as multi-tenant (for shared client-onboarding app).",
    )
    args = parser.parse_args()

    try:
        asyncio.run(_run(args.tenant, args.multi_tenant))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
