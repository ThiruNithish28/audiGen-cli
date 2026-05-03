import os

import click
import questionary
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import box

from audigen_cli import config as cfg
from audigen_cli.banner import print_banner
from audigen_cli.ui import custom_style
from audigen_cli.utils import is_word_file,_validate_date,_validate_date_range,resolve_output_dir

console = Console()

CONFIG_FIELDS = {
    "api_key": {
        "label":    "Gemini API Key",
        "prompt":   "Enter your Gemini API key",
        "password": True,
        "mask_fn":  lambda v: v[:6] + "..." + v[-4:] if v and len(v) > 10 else v,
    },
    "default_user": {
        "label":    "Default User",
        "prompt":   "Enter your name",
        "password": False,
        "mask_fn":  None,
    },
    "default_approver": {
        "label":    "Default Approver",
        "prompt":   "Enter default approver name",
        "password": False,
        "mask_fn":  None,
    },
    "output_dir": {
        "label":    "Output Directory",
        "prompt":   "Enter output folder path",
        "password": False,
        "mask_fn":  None,
    },
}

def _prompt_for_field(key: str) -> str | None:
    """Prompt the user for a single config field. Returns value or None if skipped."""
    field= CONFIG_FIELDS[key]

    if field["password"]:
        value = _ask(questionary.password(field["prompt"] + ":", style=custom_style))
    else:
        value = _ask(questionary.text(field["prompt"] + ":", style=custom_style))
    
    if not value or not value.strip():
        console.print(f"[yellow]⚠ Skipped {field['label']} (no input)[/yellow]")
        return None
    # Special handling for output_dir to validate path
    if key == "output_dir":
        path = Path(value.strip())
        if not path.exists():
            create = _ask(questionary.confirm(f"Folder does not exist. Create it?", style=custom_style))
            if create:
                try:
                    path.mkdir(parents=True, exist_ok=True)
                    console.print(f"[green]✔ Created folder:[/green] {path}")
                except Exception as e:
                    console.print(f"[red]✘ Failed to create folder:[/red] {e}")
                    return None
            else:
                console.print("[yellow]⚠ Skipped output directory.[/yellow]")
                return None
        return str(path)
    return value.strip()

def _ask(question_fn):
    """Run a questionary prompt. Exit cleanly if user presses Ctrl+C."""
    result = question_fn.ask()
    if result is None:
        console.print("\n[yellow]Aborted.[/yellow]")
        raise SystemExit(0)
    return result

# For validating date inputs if user enter the date via flag
def _assert_valid_date(value: str):
    result=_validate_date(value)
    if result is not True:
        console.print(f"[red]✘ {result}[/red]")
        raise SystemExit(1)

# ─────────────────────────────────────────────
# Root group
# ─────────────────────────────────────────────
@click.group()
def cli():
   """AudiGen — Audit document generator CLI."""
   pass
 
 
# ─────────────────────────────────────────────
# Config commands
# ─────────────────────────────────────────────
@cli.group()
def config():
    """Manage AuditGen configuration"""
    pass

@config.command("setup")
def config_setup():
    """Interactive setup — select which fields to configure."""
    choices = [
        questionary.Choice(title=meta["label"], value=key)
        for key, meta in CONFIG_FIELDS.items()
    ]
    selected_keys = questionary.checkbox(
        "Select fields to configure:(space to toggle, enter to confirm)", 
        choices=choices
    ).ask()

    if not selected_keys:
        console.print("[yellow]⚠ No fields selected. Aborting setup.[/yellow]")
        return
    
    console.print()
    for key in selected_keys:
        value = _prompt_for_field(key)
        if value:
            cfg.set_value(key, value)
            console.print(f"[green]✔[/green] {CONFIG_FIELDS[key]['label']} saved.\n")

@config.command("show")
def config_show():
    """Display current configuration"""
    current = cfg.load_config()
    table = Table(title="AuditGen Configuration", box=box.ROUNDED, show_header=True)
    table.add_column("Key", style='cyan', no_wrap=True)
    table.add_column("Value", style='white')

    for key, meta in CONFIG_FIELDS.items():
        raw = current.get(key)
        if raw and meta["mask_fn"]:
            display = meta["mask_fn"](raw)
        else:
            display = raw or "[dim] not set [/dim]"
        table.add_row(meta["label"], display)

    console.print()
    console.print(table)
    console.print("Use 'auditgen config setup' to update these values.")
    console.print()

# ─────────────────────────────────────────────
# Generate command
# ─────────────────────────────────────────────

COMPLEXITY_CHOICES = click.Choice(["LOW", "MEDIUM", "HIGH"], case_sensitive=False)
PRIORITY_CHOICES   = click.Choice(["P1", "P2", "P3"],        case_sensitive=False)

@cli.command()
@click.argument("brd",  type=click.Path(exists=True),required=False, default=None)
@click.argument("ticket",required=False, default=None)
@click.option("--start", "-s",default=None ,  help="BRD start date (DD-MM-YYYY).")
@click.option("--end","-e",default=None ,  help="BRD end date (DD-MM-YYYY).")
@click.option("--user", "-u", default=None, help="Your name. Falls back to config default.")
@click.option("--complexity", "-c", default=None,type=COMPLEXITY_CHOICES,help="Ticket complexity.")
@click.option("--priority", "-p", default=None,type=PRIORITY_CHOICES,help="Ticket priority.")
@click.option("--approver", "-a", default=None,  help="Approver name for the test case sheet.")
@click.option("--output", "-o", default=None,   help="Output folder. Falls back to config, then current directory.")
def generate(brd, ticket, start, end, user, complexity, priority, approver, output):
    print_banner()

    # ── Check API key ─────────────────────────
    api_key = cfg.get("api_key")
    if not api_key:
        console.print("[red]✘ Gemini API key not configured.[/red]")
        console.print("  Run [bold cyan]auditgen config setup[/bold cyan] first.")
        raise SystemExit(1)
 

    # ── Resolve every input — flag → prompt fallback ──────────────────
    brd = brd or _ask(questionary.path("BRD file path:", style=custom_style, validate=is_word_file,))
    if not is_word_file(brd):    # Validate if user provided date via flag
        console.print("[red]✘ BRD must be a .doc or .docx file.[/red]")
        raise SystemExit(1)
    
    ticket = ticket or _ask(questionary.text("Ticket ID:", validate=lambda text: True if len(text) > 0 else "Please enter a value", style=custom_style))

    start= start or _ask(questionary.text ("Start date (DD-MM-YYYY):", style=custom_style, validate=_validate_date))
    _assert_valid_date(start) # Validate if user provided date via flag
    
    end = end or _ask(questionary.text ("End date (DD-MM-YYYY):", style=custom_style, validate=_validate_date))
    _assert_valid_date(end) # Validate if user provided date via flag

    # Date range check 
    if not _validate_date_range(start, end):
        console.print("[red]✘ Start date cannot be after end date.[/red]")
        raise SystemExit(1)
    
    resolved_user = user or cfg.get("default_user") or _ask(questionary.text("Your name:", style=custom_style))
    
    resolved_approver = approver or cfg.get("default_approver") or _ask(questionary.text("Approver name:", style=custom_style))
    
    complexity = complexity or _ask(questionary.select("Complexity:", choices=["LOW", "MEDIUM", "HIGH"], default="MEDIUM", style=custom_style))
    
    priority = priority or _ask(questionary.select("Priority:", choices=["P1", "P2", "P3"], default="P2", style=custom_style))
    
    # ── Resolve output dir ────────────────────
    try:
        out_dir = resolve_output_dir(ticket, output)
    except ValueError as e:
        console.print(f"[red]✘ {e}[/red]")
        raise SystemExit(1)

    # ── Summary panel before running ──────────
    summary = (
        f"[cyan]BRD:[/cyan]        {brd}\n"
        f"[cyan]Ticket:[/cyan]     {ticket}\n"
        f"[cyan]Dates:[/cyan]      {start}  →  {end}\n"
        f"[cyan]User:[/cyan]       {resolved_user}\n"
        f"[cyan]Complexity:[/cyan] {complexity}   [cyan]Priority:[/cyan] {priority}\n"
        f"[cyan]Approver:[/cyan]   {resolved_approver}"
    )
    console.print(Panel(summary, title="[bold white]Generate Run[/bold white]", box=box.ROUNDED,title_align="left"))
    console.print()

    # ── Step 1: Extract BRD ───────────────────
    with console.status("[bold cyan][1/3] Extracting BRD...[/bold cyan]", spinner="dots"):
        from audigen_cli.extractor import extractDoc
        sanitized_text =extractDoc(brd)
    console.print("[green]✔[/green] [1/3] BRD extracted.")

    # ── Step 2: LLM ───────────────────────────
    os.environ["GEMINI_API_KEY"] = api_key
    with console.status("[bold cyan][2/3] Generating test cases via Gemini...[/bold cyan]", spinner="dots2"):
        from audigen_cli.llm_client import callLLM
        llm_result = callLLM(sanitized_text)  
    console.print("[green]✔[/green] [2/3] Test cases generated.")
    
    # ── Step 3: Write Excel ───────────────────
    with console.status("[bold cyan][3/3] Writing Excel files...[/bold cyan]", spinner="dots"):
        from audigen_cli.excelWriter import startExcelChange
        startExcelChange(
            llm_generateTestCase=llm_result,
            BRD_startDate=start,
            BRD_endDate=end,
            out_dir=str(out_dir),
            approver=resolved_approver,
            user=resolved_user,
            ticket=ticket,
        )
    console.print("[green]✔[/green] [3/3] Excel files written.")
    # ── Done ──────────────────────────────────
    info = llm_result.additional_info 
    console.print()
    console.print(Panel(
        f"[green]All documents generated successfully![/green]\n"
        f"[dim]Saved to:[/dim] [bold]{out_dir}[/bold]\n"
        f"[dim]Components Affected:[/dim] [bold]{info.componets_affected}[/bold]\n"  
        f"[dim]Raised By:[/dim] [bold]{info.brd_rasiedBy}[/bold]",
        box=box.ROUNDED,
        border_style="green"
    ))
    console.print()
    
if __name__ == "__main__":
    import traceback
    try:
        cli()
    except Exception as e:
        log_path = Path.home() / "auditgen_crash.log"
        with open(log_path, "w") as f:
            f.write(traceback.format_exc())
        print(f"Crashed. Log saved to: {log_path}")
        raise SystemExit(1)