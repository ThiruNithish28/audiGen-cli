from datetime import datetime
import os

import click
from pathlib import Path
from rich.prompt import Prompt, Confirm
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import box

from audigen_cli import config as cfg
from audigen_cli.banner import print_banner

console = Console()
DATE_FORMAT = "%d-%m-%Y"

def parse_date(value: str) -> str:
    """Validate and normalize date input."""
    try:
        datetime.strptime(value, DATE_FORMAT)
        return value
    except ValueError:
        raise click.BadParameter(f"Expected DD-MM-YYYY, got: {value}")

def resolve_output_dir(ticket_id: str, output_arg: str | None) -> Path:
    """
    Priority: --output arg > config default > current working directory.
    Always creates a subfolder named after the ticket.
    """
    base = output_arg or cfg.get("output_dir") or os.getcwd()
    out = Path(base) / ticket_id
    out.mkdir(parents=True, exist_ok=True)
    return out

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

@config.command("set-key")
def config_set_key():
    """set your Gemini API key."""
    key = Prompt.ask("[bold cyan]Enter your Gemini API key[/bold cyan]", password=True)
    if not key.strip():
        console.print("[red]API key cannot be empty.[/red]")
        return
    cfg.set_value("api_key", key.strip())
    console.print("[green]✔[/green] API key saved.")

@config.command("set-user")
def config_set_user():
    """Set your default username (used in generated documents)."""
    name = Prompt.ask("[bold cyan]Enter your name[/bold cyan]")
    if not name.strip():
        console.print("[red]Name cannot be empty.[/red]")
        return
    cfg.set_value("default_user", name.strip())
    console.print(f"[green]✔[/green] Default user set to [bold]{name.strip()}[/bold].")

@config.command("set-output")
def config_set_output():
    path_str = Prompt.ask("[bold cyan] Enter output folder path[/bold cyan]")
    path= Path(path_str.strip())
    if not path.exists():
        create = Confirm.ask(f"[yellow]Folder does not exist. Create it?[/yellow]")
        if create:
            path.mkdir(parents=True, exist_ok=True)
        else:
            console.print("[red]Aborted.[/red]")
            return
    cfg.set_value("output_dir", str(path))
    console.print(f"[green]✔[/green] Output folder set to [bold]{path}[/bold].")

@config.command("show")
def config_show():
    """Display current confiuration"""
    current = cfg.load_config()
    table = Table(title="AuditGen Configuration", box=box.ROUNDED, show_header=True)
    table.add_column("Key", style='cyan', no_wrap=True)
    table.add_column("Value", style='white')

    api_key = current.get('api_key')
    masked_key = (api_key[6]+"..."+api_key[-4]) if api_key and len(api_key) > 10 else ("[dim] not set [/dim]" if not api_key else api_key)
    table.add_row("api-key", masked_key)
    table.add_row("default user",current.get('default_user') or "[dim] not set [/dim]")
    table.add_row("output path", current.get("output_dir")or "[dim] not set [/dim]")

    console.print()
    console.print(table)
    console.print()

# ─────────────────────────────────────────────
# Generate command
# ─────────────────────────────────────────────

COMPLEXITY_CHOICES = click.Choice(["LOW", "MEDIUM", "HIGH"], case_sensitive=False)
PRIORITY_CHOICES   = click.Choice(["P1", "P2", "P3"],        case_sensitive=False)

@cli.command()
@click.option("--brd", "-b", required=True, type=click.Path(exists=True), help="Path to the BRD .docx file.")
@click.option("--ticket","-t", required=True,help="Ticket /issue Id")
@click.option("--start", "-s", required=True,  help="BRD start date (DD-MM-YYYY).")
@click.option("--end","-e", required=True,  help="BRD end date (DD-MM-YYYY).")
@click.option("--user", "-u", default=None, help="Your name. Falls back to config default.")
@click.option("--complexity", "-c", default="MEDIUM", type=COMPLEXITY_CHOICES, show_default=True, help="Ticket complexity.")
@click.option("--priority", "-p", default="P2",    type=PRIORITY_CHOICES,    show_default=True, help="Ticket priority.")
@click.option("--approver", "-a", required=True,  help="Approver name for the test case sheet.")
@click.option("--output", "-o", default=None,   help="Output folder. Falls back to config, then current directory.")
def generate(brd, ticket, start, end, user, complexity, priority, approver, output):
    print_banner()
     # ── Validate dates ───────────────────────
    try:
        start=parse_date(start)
        end=parse_date(end)
    except click.BadParameter as e:
        console.print(f"[red]✘ Date error:[/red] {e}")
        raise SystemExit(1)
    
    if datetime.strptime(start, DATE_FORMAT) > datetime.strptime(end, DATE_FORMAT):
        console.print("[red]✘ Start date cannot be after end date.[/red]")
        raise SystemExit(1)
    
    # ── Resolve user ─────────────────────────
    resolved_user = user or cfg.get("default_user")
    if not resolved_user:
        resolved_user=Prompt.ask("[yellow]No default user set. Enter your name[/yellow]")
    
    # ── Check API key ─────────────────────────
    api_key = cfg.get("api_key")
    if not api_key:
        console.print("[red]✘ Gemini API key not configured.[/red]")
        console.print("  Run [bold cyan]auditgen config set-key[/bold cyan] first.")
        raise SystemExit(1)
 
    os.environ["GEMINI_API_KEY"] = api_key
    # ── Summary panel before running ──────────
    summary = (
        f"[cyan]BRD:[/cyan]        {brd}\n"
        f"[cyan]Ticket:[/cyan]     {ticket}\n"
        f"[cyan]Dates:[/cyan]      {start}  →  {end}\n"
        f"[cyan]User:[/cyan]       {resolved_user}\n"
        f"[cyan]Complexity:[/cyan] {complexity}   [cyan]Priority:[/cyan] {priority}\n"
        f"[cyan]Approver:[/cyan]   {approver}"
    )
    console.print(Panel(summary, title="[bold white]Generate Run[/bold white]", box=box.ROUNDED))
    console.print()

    # ── Resolve output dir ────────────────────
    out_dir = resolve_output_dir(ticket, output)

    # ── Step 1: Extract BRD ───────────────────
    with console.status("[bold cyan][1/3] Extracting BRD...[/bold cyan]", spinner="dots"):
        from audigen_cli.extractor import extractDoc
        sanitized_text =extractDoc(brd)
    console.print("[green]✔[/green] [1/3] BRD extracted.")

    # ── Step 2: LLM ───────────────────────────
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
            approver=approver,
            user=resolved_user,
            ticket=ticket,
        )
    console.print("[green]✔[/green] [3/3] Excel files written.")
    # ── Done ──────────────────────────────────
    console.print()
    console.print(Panel(
        f"[green]All documents generated successfully![/green]\n"
        f"[dim]Saved to:[/dim] [bold]{out_dir}[/bold]",
        box=box.ROUNDED,
        border_style="green"
    ))
    console.print()