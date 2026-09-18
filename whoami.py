"""
whoami.py — Print the Asana IDs you need for .env.

    python whoami.py
"""
from __future__ import annotations

import os

from dotenv import load_dotenv

from asana_client import AsanaClient, AsanaError


def main() -> None:
    load_dotenv()
    client = AsanaClient(os.getenv("ASANA_TOKEN", ""))
    me = client.get("/users/me", {"opt_fields": "name,email,workspaces.name"})
    print(f"\nAuthenticated as {me['name']} <{me.get('email','')}>\n")
    print("Workspaces (copy the gid of the one your firm uses into ASANA_WORKSPACE_GID):")
    for ws in me.get("workspaces", []):
        print(f"  {ws['gid']}  {ws['name']}")
        try:
            teams = list(client.paginate(f"/organizations/{ws['gid']}/teams",
                                         {"opt_fields": "name"}))
            for t in teams[:15]:
                print(f"      team  {t['gid']}  {t['name']}")
            if len(teams) > 15:
                print(f"      ... and {len(teams)-15} more teams")
        except AsanaError:
            print("      (not an organization, or no permission to list teams)")
    print()


if __name__ == "__main__":
    try:
        main()
    except AsanaError as exc:
        raise SystemExit(f"Asana error: {exc}")
