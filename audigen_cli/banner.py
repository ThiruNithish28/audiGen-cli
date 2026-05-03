from rich.console import Console
from rich.text import Text
from rich.console import Console
from rich.panel import Panel
from rich.align import Align

console =Console()

VERSION = "0.1.0"

ASCII_ART = """\
 █████╗ ██╗   ██╗██████╗ ██╗████████╗ ██████╗ ███████╗███╗   ██╗
██╔══██╗██║   ██║██╔══██╗██║╚══██╔══╝██╔════╝ ██╔════╝████╗  ██║
███████║██║   ██║██║  ██║██║   ██║   ██║  ███╗█████╗  ██╔██╗ ██║
██╔══██║██║   ██║██║  ██║██║   ██║   ██║   ██║██╔══╝  ██║╚██╗██║
██║  ██║╚██████╔╝██████╔╝██║   ██║   ╚██████╔╝███████╗██║ ╚████║
╚═╝  ╚═╝ ╚═════╝ ╚═════╝ ╚═╝   ╚═╝    ╚═════╝ ╚══════╝╚═╝  ╚═══╝"""

def print_banner():
    art= Text(ASCII_ART, style="bold cyan")
    subtitle=Text(f"v{VERSION} | Audit Document Generator | by Thiru", style="dim white")
    console.print()
    console.print(art)
    console.print(subtitle)
    console.print()


def print_banner2():
    # Brand line
    title = Text()
    title.append("AuditGen", style="bold #ff8c69")
    title.append(f"  v{VERSION}", style="bold white")

    # Body
    body = Text()
    body.append("AI Audit Document Generator\n", style="bold white")
    body.append("Generate test cases, review outputs, export Excel.\n", style="white")
    body.append("\n")
    body.append("/generate", style="bold #7aa2f7")
    body.append("   Create audit documents from BRD\n", style="dim white")
    body.append("/config", style="bold #7aa2f7")
    body.append("     Manage API key, user, output folder\n", style="dim white")
    body.append("/review", style="bold #7aa2f7")
    body.append("     Inspect latest generated run\n", style="dim white")
    body.append("/help", style="bold #7aa2f7")
    body.append("       Show available commands", style="dim white")

    panel = Panel(
        Align.left(body),
        title=title,
        title_align="left",
        border_style="#ff8c69",
        padding=(1, 2),
        expand=True,
    )

    console.print()
    console.print(panel)
    console.print("[dim]Ready. Type a command to continue.[/dim]")
    console.print()

