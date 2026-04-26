"""`outlook skill` sub-app — manage the OpenClaw skill bundle that ships with the CLI."""

from __future__ import annotations

import os
import shutil
from importlib.resources import as_file, files
from pathlib import Path
from typing import Annotated

import typer

from outlook_cli.commands._common import err_console

app = typer.Typer(help="OpenClaw skill management.", no_args_is_help=True)

# Path inside the package where bundled skill files live (force-included via pyproject).
_BUNDLED_PARTS = ("_skills", "outlook")


def _resolve_target(workspace: Path | None, target: Path | None) -> Path:
    """Pick the install destination based on flags, env, and defaults.

    Priority:
      1. --target <path>           — raw destination (used as-is)
      2. --workspace <path>        — installs to <path>/skills/outlook
      3. OPENCLAW_WORKSPACE env    — same as --workspace
      4. Default                   — ~/.openclaw/skills/outlook  (managed/local)
    """
    if target is not None and workspace is not None:
        err_console.print("[red]--target and --workspace are mutually exclusive.[/red]")
        raise typer.Exit(2)

    if target is not None:
        return target.expanduser().resolve()

    if workspace is None:
        env_ws = os.environ.get("OPENCLAW_WORKSPACE")
        if env_ws:
            workspace = Path(env_ws)

    if workspace is not None:
        return workspace.expanduser().resolve() / "skills" / "outlook"

    return Path.home() / ".openclaw" / "skills" / "outlook"


@app.command("install")
def install(
    workspace: Annotated[
        Path | None,
        typer.Option(
            "--workspace",
            help=(
                "Install to <workspace>/skills/outlook (workspace-level, highest "
                "precedence in OpenClaw). Falls back to OPENCLAW_WORKSPACE env if unset."
            ),
        ),
    ] = None,
    target: Annotated[
        Path | None,
        typer.Option(
            "--target",
            help=(
                "Install to this exact path (overrides --workspace). Use for "
                "non-standard layouts e.g. <workspace>/.agents/skills/outlook."
            ),
        ),
    ] = None,
    force: Annotated[
        bool,
        typer.Option("--force", help="Overwrite an existing installation."),
    ] = False,
):
    """Install the bundled OpenClaw skill onto disk.

    Resolution order:
      1. --target <path>            install to <path> exactly
      2. --workspace <path>         install to <path>/skills/outlook
      3. $OPENCLAW_WORKSPACE        install to $OPENCLAW_WORKSPACE/skills/outlook
      4. (default)                  install to ~/.openclaw/skills/outlook

    Common OpenClaw skill locations (precedence: highest first):
      <workspace>/skills/outlook                    workspace skills (highest)
      <workspace>/.agents/skills/outlook            project agent skills
      ~/.agents/skills/outlook                      personal agent skills
      ~/.openclaw/skills/outlook                    managed/local (default)
    """
    dest = _resolve_target(workspace, target)

    if dest.exists() and not force:
        err_console.print(
            f"[yellow]Skill already installed at {dest}.[/yellow]\n"
            "Use --force to overwrite, or `outlook skill uninstall` first."
        )
        raise typer.Exit(1)

    src = files("outlook_cli")
    for part in _BUNDLED_PARTS:
        src = src / part

    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        shutil.rmtree(dest)

    with as_file(src) as src_path:
        if not src_path.exists():
            err_console.print(
                f"[red]Bundled skill files not found at {src_path}.[/red]\n"
                "This usually means the package was installed from source without "
                "running the wheel build. Reinstall via "
                "`uv tool install --upgrade git+https://github.com/alavida-ai/outlook-cli`."
            )
            raise typer.Exit(2)
        shutil.copytree(src_path, dest)

    err_console.print(f"[green]Skill installed at[/green] {dest}")
    err_console.print(
        "Restart OpenClaw to pick it up:  [cyan]openclaw gateway restart[/cyan]"
    )
    err_console.print(
        "Then verify:                    [cyan]openclaw skills list[/cyan]"
    )


@app.command("uninstall")
def uninstall(
    workspace: Annotated[
        Path | None,
        typer.Option(
            "--workspace",
            help="Uninstall from <workspace>/skills/outlook. Honours OPENCLAW_WORKSPACE.",
        ),
    ] = None,
    target: Annotated[
        Path | None,
        typer.Option("--target", help="Uninstall from this exact path."),
    ] = None,
):
    """Remove the OpenClaw skill from disk."""
    dest = _resolve_target(workspace, target)

    if not dest.exists():
        err_console.print(f"[yellow]No skill at {dest}.[/yellow]")
        raise typer.Exit(1)

    shutil.rmtree(dest)
    err_console.print(f"[green]Skill removed from[/green] {dest}")


@app.command("path")
def path(
    bundled: Annotated[
        bool,
        typer.Option(
            "--bundled/--installed",
            help="Show the bundled source path (default) or the resolved install path.",
        ),
    ] = True,
    workspace: Annotated[
        Path | None,
        typer.Option("--workspace", help="See `outlook skill install --help` for resolution."),
    ] = None,
    target: Annotated[
        Path | None, typer.Option("--target", help="Raw path override.")
    ] = None,
):
    """Print a path:
      --bundled (default)  the bundled skill source inside the installed wheel
      --installed          the on-disk install destination (uses same resolution as `install`)
    """
    if bundled:
        src = files("outlook_cli")
        for part in _BUNDLED_PARTS:
            src = src / part
        with as_file(src) as src_path:
            err_console.print(str(src_path))
        return

    dest = _resolve_target(workspace, target)
    err_console.print(str(dest))
