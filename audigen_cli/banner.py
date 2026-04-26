from rich.console import Console
from rich.text import Text

console =Console()

VERSION = "0.1.0"

ASCII_ART = """\
 █████╗ ██╗   ██╗██████╗ ██╗ ██████╗ ███████╗███╗   ██╗
██╔══██╗██║   ██║██╔══██╗██║██╔════╝ ██╔════╝████╗  ██║
███████║██║   ██║██║  ██║██║██║  ███╗█████╗  ██╔██╗ ██║
██╔══██║██║   ██║██║  ██║██║██║   ██║██╔══╝  ██║╚████║
██║  ██║╚██████╔╝██████╔╝██║╚██████╔╝███████╗██║  ╚███║
╚═╝  ╚═╝ ╚═════╝ ╚═════╝ ╚═╝ ╚═════╝ ╚══════╝╚═╝   ╚══╝"""

def print_banner():
    art= Text(ASCII_ART, style="bold cyan")
    subtitle=Text(f"v{VERSION} | Audit Document Generator | by Thiru", style="dim white")
    console.print()
    console.print(art)
    console.print(subtitle)
    console.print()
