from questionary import Style

custom_style = Style([
    ("qmark", "fg:#00d7ff bold"),
    ("question", "fg:#00d7ff bold"),

    # Input answers (text prompts)
    ("answer", "fg:#ffffff bold"),

    # Select-specific
    ("pointer", "fg:#00d7ff bold"),       # ❯ arrow
    ("highlighted", "fg:#00d7ff bold"),   # active option
    ("selected", "fg:#00d7ff bold"),      # after selection

    # Secondary UI
    ("separator", "fg:#6c6c6c"),
    ("instruction", "fg:#6c6c6c"),
])