"""Microsoft Graph client factory."""

from __future__ import annotations

from azure.core.credentials import AccessToken, TokenCredential
from msgraph import GraphServiceClient

from outlook_cli import auth


class _MsalTokenCredential(TokenCredential):
    """TokenCredential delegating to MSAL on every call so refresh works in long-running clients."""

    def __init__(self, tenant_id: str, client_id: str) -> None:
        self._tenant_id = tenant_id
        self._client_id = client_id

    def get_token(self, *scopes: str, **kwargs) -> AccessToken:  # noqa: D401 - SDK signature
        token, expires_on = auth.get_access_token(self._tenant_id, self._client_id)
        return AccessToken(token, expires_on=expires_on)


def get_client(tenant_id: str, client_id: str) -> GraphServiceClient:
    """Return a GraphServiceClient that auto-refreshes via MSAL on each call."""
    return GraphServiceClient(credentials=_MsalTokenCredential(tenant_id, client_id))