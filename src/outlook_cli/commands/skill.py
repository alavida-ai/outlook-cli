"""`outlook skill` sub-app — manage the OpenClaw skill bundle that ships with the CLI."""

from __future__ import annotations

import shutil
from importlib.resources import as_file, files
from pathlib import Path
from typing import Annotated

import typer

from outlook_cli.commands._common import err_console

app = typer.Typer(help="OpenClaw skill management.", no_args_is_help=True)

# Path inside the package where bundled skill files live (force-included via pyproject).
_BUNDLED_PARTS = ("_skills", "outlook")


def _resolve_target(workspace: Path | None) -> Path:
    if workspace is not None:
        return workspace.expanduser().resolve() / "skills" / "outlook"
    return Path.home() / ".openclaw" / "skills" / "outlook"


@app.command("install")
def install(
    workspace: Annotated[
        Path | None,
        typer.Option(
            "--workspace",
            help="Install into <workspace>/skills/outlook instead of ~/.openclaw/skills/outlook.",
        ),
    ] = None,
    force: Annotated[
        bool,
        typer.Option("--force", help="Overwrite an existing installation."),
    ] = False,
):
    """Install the bundled OpenClaw skill (default: ~/.openclaw/skills/outlook).

    The skill files (SKILL.md + references/*) ship inside the outlook-cli wheel.
    This command copies them onto disk where OpenClaw's gateway scans for skills.
    """
    target = _resolve_target(workspace)

    if target.exists() and not force:
        err_console.print(
            f"[yellow]Skill already installed at {target}.[/yellow]\n"
            "Use --force to overwrite, or `outlook skill uninstall` first."
        )
        raise typer.Exit(1)

    src = files("outlook_cli")
    for part in _BUNDLED_PARTS:
        src = src / part

    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        shutil.rmtree(target)

    with as_file(src) as src_path:
        if not src_path.exists():
            err_console.print(
                f"[red]Bundled skill files not found at {src_path}.[/red]\n"
                "This usually means the package was installed from source without "
                "running the wheel build. Reinstall via "
                "`uv tool install --upgrade git+https://github.com/alavida-ai/outlook-cli`."
            )
            raise typer.Exit(2)
        shutil.copytree(src_path, target)

    err_console.print(f"[green]Skill installed at[/green] {target}")
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
            help="Uninstall from <workspace>/skills/outlook instead of ~/.openclaw/skills/outlook.",
        ),
    ] = None,
):
    """Remove the OpenClaw skill from disk."""
    target = _resolve_target(workspace)

    if not target.exists():
        err_console.print(f"[yellow]No skill at {target}.[/yellow]")
        raise typer.Exit(1)

    shutil.rmtree(target)
    err_console.print(f"[green]Skill removed from[/green] {target}")


@app.command("path")
def path():
    """Show where the bundled skill source lives inside the installed package (for inspection)."""
    src = files("outlook_cli")
    for part in _BUNDLED_PARTS:
        src = src / part
    with as_file(src) as src_path:
        err_console.print(str(src_path))
