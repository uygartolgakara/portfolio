from PIL import Image, ImageDraw, ImageFont

# LinkedIn banner dimensions
WIDTH, HEIGHT = 1584, 396
BACKGROUND_COLOR = (40, 44, 52)  # Dark background similar to Spyder
image = Image.new('RGB', (WIDTH, HEIGHT), BACKGROUND_COLOR)
draw = ImageDraw.Draw(image)

# Font settings with fallback
try:
    font = ImageFont.truetype("consola.ttf", 24)  # Windows
except IOError:
    try:
        font = ImageFont.truetype("Menlo.ttc", 24)  # Mac
    except IOError:
        font = ImageFont.load_default()  # Fallback

# Syntax highlighting colors
COLORS = {
    'keyword': (197, 165, 197),  # Purple
    'def': (86, 182, 194),      # Cyan
    'string': (152, 195, 121),  # Green
    'comment': (106, 153, 85),  # Gray-green
    'normal': (220, 220, 220),  # Light gray
    'class': (239, 200, 145),   # Orange
}

# Code content with improved spacing
code_lines = [
    ("class ", 'keyword'), ("LinkedInProfile", 'class'), (":", 'normal'), ("\n", None),
    ("    ", None), ("def ", 'keyword'), ("__init__", 'def'), ("(self):", 'normal'), ("\n", None),
    ("        ", None), ("self.name = ", 'normal'), ('"Uygar Tolga Kara"', 'string'), ("\n", None),
    ("        ", None), ("self.title = ", 'normal'), ('"Python Developer | Automation Expert | Data Specialist"', 'string'), ("\n", None),
    ("        ", None), ("self.tech_stack = ", 'normal'), ("[", 'normal'),
    ('"Python"', 'string'), (", ", 'normal'),
    ('"VBA"', 'string'), (", ", 'normal'),
    # ('"Power Automate"', 'string'), (", ", 'normal'),
    # ('"UiPath"', 'string'), (", ", 'normal'),
    ('"Power BI"', 'string'), (", ", 'normal'),
    ('"Tableau"', 'string'), (", ", 'normal'),
    ('"Azure"', 'string'), ("]", 'normal'), ("\n", None),
    # ("        ", None), ("self.stats = ", 'normal'), ("{", 'normal'),
    # ('"years_exp"', 'string'), (": ", 'normal'), ('"5+"', 'string'), (", ", 'normal'),
    # ('"projects"', 'string'), (": ", 'normal'), ('"100+"', 'string'), ("}", 'normal'), ("\n", None),
    ("\n", None),  # Extra spacing before comment
    ("            ", None), ("# ", 'comment'), ("Open to collaboration on automation & data projects", 'comment')
]

# Text positioning
x, y = 50, 20
line_height = 30
tab_size = 4 * 24  # Approximate width of 4 spaces in the font

# Rendering engine
for text, color_type in code_lines:
    if text == "\n":
        y += line_height
        x = 50
        continue
    if text.startswith("    "):
        x += tab_size * text.count("    ")
        text = text.lstrip("    ")

    color = COLORS[color_type] if color_type else COLORS['normal']
    draw.text((x, y), text, font=font, fill=color)
    x += draw.textlength(text, font=font)

# Final output
image.save("linkedin_banner_code_style.png")
image.show()