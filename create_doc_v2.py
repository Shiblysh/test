from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_hyperlink(paragraph, text, url):
    """
    A function that creates a hyperlink inside a paragraph using python-docx.
    Properly handles the relationship and XML structure.
    """
    part = paragraph.part
    
    # Create the relationship for the hyperlink
    r_id = part.relate_to(url, 'external')
    
    # Create the w:hyperlink element
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    hyperlink.set(qn('w:history'), '1')  # Mark as visited/history
    
    # Create the run element with styling
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    # Add styles for hyperlink (blue color and underline)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')
    color.set(qn('w:themeColor'), 'hyperlink')
    rPr.append(color)
    
    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    rPr.append(underline)
    
    new_run.append(rPr)
    
    # Add text
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    
    # Append run to hyperlink
    hyperlink.append(new_run)
    
    # Append hyperlink to paragraph's element
    paragraph._element.append(hyperlink)
    
    return hyperlink

# Create document
doc = Document()

# Title
title = doc.add_heading('The Ultimate Guide to Eco Friendly Cutting Boards: Sustainable Choices for a Healthier Kitchen', 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Meta Description
p = doc.add_paragraph()
run = p.add_run('Meta Description: ')
run.bold = True
p.add_run('Discover the best eco friendly cutting board options for your kitchen. Learn about sustainable materials like bamboo and FSC-certified wood, avoid greenwashing, and master care tips for longevity. Switch to a healthier, planet-friendly kitchen today.')
p.style = 'Intense Quote'

# Section I
doc.add_heading('I. Introduction', level=1)

p = doc.add_paragraph()
run = p.add_run('Hook: ')
run.bold = True
p.add_run('Every time you chop vegetables or slice meat on a plastic cutting board, you might be contributing to a hidden environmental crisis. Microplastics are shedding into your food and the environment, while discarded boards pile up in landfills, taking centuries to decompose.')

p = doc.add_paragraph()
run = p.add_run('Problem: ')
run.bold = True
p.add_run("Most home cooks don't realize their cutting board affects both their health and the planet. Plastic boards harbor bacteria in deep knife scars and contribute to global plastic pollution.")

p = doc.add_paragraph()
run = p.add_run('Solution: ')
run.bold = True
p.add_run('The rise of the ')
add_hyperlink(p, 'eco friendly cutting board', 'https://bestecokitchentools.com/')
p.add_run(' offers a sustainable, healthy, and high-performance alternative.')

p = doc.add_paragraph()
run = p.add_run('Thesis: ')
run.bold = True
p.add_run('This guide covers the best sustainable materials, what to avoid, and how to care for your board to ensure it lasts a lifetime.')

# Section II
doc.add_heading('II. Why Choose an Eco Friendly Cutting Board?', level=1)

doc.add_heading('Health Benefits', level=2)
p = doc.add_paragraph()
p.add_run('Unlike plastic, which develops deep grooves where bacteria like Salmonella and E. coli thrive, natural materials like wood and bamboo possess inherent antimicrobial properties. Studies have shown that wood surfaces can actually kill bacteria over time, making them a safer choice for your family\'s health.')

doc.add_heading('Environmental Impact', level=2)
p = doc.add_paragraph()
p.add_run('Choosing an eco friendly cutting board means supporting biodegradability and renewable resources. Bamboo grows back in just 3-5 years, and responsibly harvested hardwoods ensure forests remain intact. By avoiding plastic, you reduce the demand for petroleum-based products and help curb plastic pollution in our oceans and landfills.')

doc.add_heading('Performance', level=2)
p = doc.add_paragraph()
p.add_run("Sustainable boards aren't just good for the planet; they perform better too. Wood and bamboo are gentler on knife edges, keeping your blades sharper for longer. Plus, the natural aesthetic appeal of grain patterns adds warmth and beauty to your kitchen countertop.")

# Section III
doc.add_heading('III. Best Materials for an Eco Friendly Cutting Board', level=1)

doc.add_heading('A. Bamboo', level=2)
p = doc.add_paragraph()
run = p.add_run("Why it's sustainable: ")
run.bold = True
p.add_run('Bamboo is technically a grass, not a tree, and is one of the fastest-growing plants on Earth. It requires no pesticides and very little water to thrive.')

p = doc.add_paragraph()
run = p.add_run('Pros: ')
run.bold = True
p.add_run('Highly durable, affordable, and naturally antimicrobial.')
run = p.add_run('\nCons: ')
run.bold = True
p.add_run('Can be harder than some woods, potentially dulling knives faster if not maintained properly.')

p = doc.add_paragraph()
p.add_run('Read more: ')
add_hyperlink(p, 'Is Bamboo a Good Wood for Cutting Boards?', 'https://bestecokitchentools.com/is-bamboo-a-good-wood-for-cutting-boards/')

doc.add_heading('B. FSC-Certified Hardwoods (Maple, Walnut, Cherry)', level=2)
p = doc.add_paragraph()
run = p.add_run('Importance of FSC certification: ')
run.bold = True
p.add_run('The Forest Stewardship Council (FSC) ensures that wood is harvested from responsibly managed forests, preventing deforestation and protecting biodiversity.')

p = doc.add_paragraph()
run = p.add_run('Pros: ')
run.bold = True
p.add_run('Knife-friendly surfaces, incredibly long-lasting, and beautiful grain patterns.')
run = p.add_run('\nCons: ')
run.bold = True
p.add_run('Requires regular oiling and hand washing to maintain integrity.')

p = doc.add_paragraph()
p.add_run('Read more: ')
add_hyperlink(p, 'What Is the Best Wood for Cutting Boards?', 'https://bestecokitchentools.com/what-is-the-best-wood-for-cutting-boards/')

doc.add_heading('C. Reclaimed Wood', level=2)
p = doc.add_paragraph()
p.add_run('Upcycling old timber from barns, factories, or fallen trees gives new life to materials that would otherwise go to waste. This option boasts unique aesthetics with character marks and history, representing zero new resource extraction.')

doc.add_heading('D. Recycled Composite (Paper/Resin)', level=2)
p = doc.add_paragraph()
p.add_run('Made from post-consumer waste like paper and non-petroleum resins, these boards are a fantastic dishwasher-safe alternative for convenience seekers. They are dense, durable, and environmentally conscious.')

# Section IV
doc.add_heading('IV. What to Avoid: Greenwashing Red Flags', level=1)

p = doc.add_paragraph()
p.add_run('Not all "green" products are truly sustainable. Watch out for these red flags:')

items = [
    ('Uncertified tropical woods (teak/acacia without FSC): ', 'These often come from illegal logging operations that destroy rainforests.'),
    ('Plastic-wood hybrids: ', 'Often non-recyclable and difficult to dispose of responsibly.'),
    ('Synthetic chemical finishes: ', 'Look for boards finished with food-grade mineral oil or beeswax, not varnishes containing VOCs.')
]

for bold_text, normal_text in items:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(bold_text)
    run.bold = True
    p.add_run(normal_text)

# Section V
doc.add_heading('V. Care and Maintenance for Longevity', level=1)

p = doc.add_paragraph()
p.add_run('To get the most out of your eco friendly cutting board, follow these simple care tips:')

care_items = [
    ('Hand washing only: ', 'Never put wood or bamboo in the dishwasher; the heat and water will warp and crack them.'),
    ('Oiling regularly: ', 'Apply food-grade mineral oil or beeswax once a month (or when the wood looks dry) to keep it hydrated.'),
    ('Natural sanitizing: ', 'Use white vinegar, lemon juice, and coarse salt to clean and deodorize naturally.'),
    ('Proper storage: ', 'Store upright in a dry area to allow air circulation and prevent moisture buildup.')
]

for bold_text, normal_text in care_items:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(bold_text)
    run.bold = True
    p.add_run(normal_text)

# Section VI
doc.add_heading('VI. FAQ Section', level=1)

doc.add_heading('Is bamboo better than wood for the environment?', level=2)
p = doc.add_paragraph()
p.add_run('Generally, yes. Bamboo grows much faster than hardwood trees (3-5 years vs. decades), making it a highly renewable resource. However, FSC-certified hardwoods are also excellent sustainable choices if harvested responsibly.')

doc.add_heading('Can I use wood boards for raw meat?', level=2)
p = doc.add_paragraph()
p.add_run("Yes, provided you clean them properly immediately after use. The antimicrobial properties of wood make it safe for meat, but you must wash with hot soapy water and sanitize with vinegar or a diluted bleach solution. Many people prefer using separate boards for produce and meat to be extra safe.")

doc.add_heading('How often should I oil my board?', level=2)
p = doc.add_paragraph()
p.add_run("It depends on usage and climate, but a good rule of thumb is once a month. If water no longer beads up on the surface but soaks in immediately, it's time to oil.")

# Section VII
doc.add_heading('VII. Conclusion', level=1)

p = doc.add_paragraph()
p.add_run('Making the switch to an ')
add_hyperlink(p, 'eco friendly cutting board', 'https://bestecokitchentools.com/')
p.add_run(' is a small change with a big impact. Whether you choose fast-growing bamboo, FSC-certified hardwood, reclaimed timber, or recycled composite, you are investing in a healthier kitchen and a cleaner planet.')

p = doc.add_paragraph()
p.add_run('Recap of top choices: Bamboo for affordability and renewability, Maple/Walnut for knife care and longevity, Reclaimed Wood for unique character, and Recycled Composite for convenience.')

p = doc.add_paragraph()
p.add_run('Ready to upgrade your kitchen? ')
add_hyperlink(p, 'Browse our Sustainable Kitchen Tools', 'https://bestecokitchentools.com/collections/sustainable-kitchen-tools')
p.add_run(' or dive deeper into our ')
add_hyperlink(p, 'Bamboo vs. Wood Guide', 'https://bestecokitchentools.com/bamboo-vs-wood-guide')
p.add_run('.')

# Save
file_path = '/workspace/Eco_Friendly_Cutting_Board_Guide.docx'
doc.save(file_path)
print(f"Document saved successfully to {file_path}")
