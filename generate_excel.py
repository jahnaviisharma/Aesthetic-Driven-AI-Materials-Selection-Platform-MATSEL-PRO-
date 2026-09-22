"""
Generate Material Database Excel workbook for CP302 IIT Ropar
Tabs: All Materials, Metals, Polymers, Natural Materials, Composites, Processes
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Color scheme ──
NAVY = "0B1120"
BLUE = "2563EB"
WHITE = "FFFFFF"
LIGHT_GREY = "F3F4F6"
BORDER_CLR = "D1D5DB"

FAMILY_COLORS = {
    "Metals":           "E0E7FF",
    "Polymers":         "FCE7F3",
    "Natural Materials":"D1FAE5",
    "Composites":       "FEF3C7",
}

FAMILY_HEADER_COLORS = {
    "Metals":           "3730A3",
    "Polymers":         "9D174D",
    "Natural Materials":"065F46",
    "Composites":       "92400E",
}

# ── Full material database (mirrors index.html) ──
MATERIALS = [
    {
        "id": 1, "family": "Metals",
        "name": "Aluminum Alloy", "grade": "AA 6063-T5",
        "surface": "Bead Blasted + Type II Clear Anodize",
        "Ra": "0.8–1.6 µm", "GU": "10–25 GU",
        "finish": "Matte/Bead-blasted",
        "tags": "Modern, Technical, Understated, Clean",
        "products": "Consumer Electronics, Wearables / Watches, Industrial Tools, Sporting Goods",
        "applications": "Luxury / Premium, High-Performance / Engineered, Office / Workspace",
        "rationale": "Excellent strength-to-weight ratio with a durable 10–25 µm oxide layer that accepts dye. Bead-blasted substrate scatters light isotropically, hiding fingerprints.",
        "sources": "ASM Handbook Vol.5 [1], ISO 4287 [2], Ashby & Johnson 2014 [3]",
        "processes": "Bead Blasting; Type II Anodizing; Type III Hard Anodize"
    },
    {
        "id": 2, "family": "Metals",
        "name": "Stainless Steel", "grade": "316L (Medical Grade)",
        "surface": "Mirror Polished / Electropolished",
        "Ra": "0.025–0.1 µm", "GU": "85–98 GU",
        "finish": "Mirror/High-gloss",
        "tags": "Premium, Pristine, Clinical, Expensive",
        "products": "Wearables / Watches, Medical Devices, Kitchenware / Cutlery, Automotive / Transport",
        "applications": "Luxury / Premium, Medical / Hygienic, High-Performance / Engineered",
        "rationale": "Highly corrosion-resistant austenitic steel with non-reactive Mo-bearing composition. Electropolished surface eliminates micro-crevices for aseptic requirements.",
        "sources": "ASM Handbook Vol.5 [1], ISO 4287 [2], CES EduPack [4]",
        "processes": "Electropolishing; Fine Lapping"
    },
    {
        "id": 3, "family": "Metals",
        "name": "Stainless Steel", "grade": "AISI 304 (18/8)",
        "surface": "Brushed No. 4 Finish",
        "Ra": "0.4–0.8 µm", "GU": "25–45 GU",
        "finish": "Brushed/Directional",
        "tags": "Professional, Directional, Durable, Industrial-refined",
        "products": "Consumer Electronics, Furniture, Architectural Hardware, Kitchenware / Cutlery",
        "applications": "Office / Workspace, Luxury / Premium, Casual / Home Use",
        "rationale": "Industry standard for appliance and architectural panels. Unidirectional grain hides minor scratches while maintaining a sophisticated metallic sheen.",
        "sources": "ASM Handbook Vol.5 [1], ISO 4287 [2], Ashby & Johnson 2014 [3]",
        "processes": "Abrasive Belt Brushing; Passivation (Citric Acid)"
    },
    {
        "id": 4, "family": "Natural Materials",
        "name": "Black Walnut", "grade": "FAS Grade Solid Hardwood",
        "surface": "Hand-sanded 320 grit + Tung Oil / Beeswax",
        "Ra": "1.6–3.2 µm", "GU": "8–20 GU",
        "finish": "Matte/Bead-blasted",
        "tags": "Warm, Heritage, Organic, Sustainable, Crafted",
        "products": "Furniture, Architectural Hardware, Consumer Electronics, Kitchenware / Cutlery",
        "applications": "Sustainable / Eco-conscious, Luxury / Premium, Casual / Home Use, Office / Workspace",
        "rationale": "Tactile warmth and visual richness from natural grain. Oil-wax finish protects against moisture while allowing the wood to develop a unique patina.",
        "sources": "Ashby & Johnson 2014 [3], CES EduPack [4], Karana et al. 2009 [6]",
        "processes": "Precision Sanding (Multi-grit); Hand Oil / Wax Application"
    },
    {
        "id": 5, "family": "Metals",
        "name": "Titanium Alloy", "grade": "Ti-6Al-4V (Grade 5)",
        "surface": "PVD Coated (TiN — Titanium Nitride)",
        "Ra": "0.15–0.4 µm", "GU": "60–85 GU",
        "finish": "Satin/Semi-gloss",
        "tags": "High-tech, Aggressive, Gold-tone, Low-friction",
        "products": "Wearables / Watches, Industrial Tools, Automotive / Transport, Medical Devices",
        "applications": "High-Performance / Engineered, Outdoor / Harsh Environment, Luxury / Premium",
        "rationale": "PVD TiN adds 1–5 µm coating with hardness exceeding 2000 HV and a distinctive gold-metallic hue. Widely used in premium watch cases and cutting tools.",
        "sources": "ASM Handbook Vol.5 [1], ISO 4287 [2], Custompart.net [5]",
        "processes": "PVD Coating (TiN/TiAlN/CrN); Precision Grinding"
    },
    {
        "id": 6, "family": "Polymers",
        "name": "Polycarbonate (PC)", "grade": "Optical / General Purpose Grade",
        "surface": "Injection Molded — SPI A-1 High Gloss",
        "Ra": "0.012–0.05 µm", "GU": "92–100 GU",
        "finish": "Mirror/High-gloss",
        "tags": "Transparent, Futuristic, Lightweight-feel, Pure",
        "products": "Consumer Electronics, Toys / Recreational, Medical Devices, Automotive / Transport, Wearables / Watches",
        "applications": "Casual / Home Use, Office / Workspace, Luxury / Premium",
        "rationale": "Impact resistance 250× greater than glass with 88% light transmission. Diamond-polished mold cavities produce near-optical surfaces directly from the tool.",
        "sources": "CES EduPack [4], Ashby & Johnson 2014 [3], Custompart.net [5]",
        "processes": "Precision Injection Molding (SPI A-1); Vapor Polishing (Acetone/DCM)"
    },
    {
        "id": 7, "family": "Composites",
        "name": "Carbon Fiber (CFRP)", "grade": "3K Twill Weave / Epoxy Matrix",
        "surface": "Vacuum Bagged + Matte Clear Coat",
        "Ra": "0.4–1.2 µm", "GU": "5–15 GU",
        "finish": "Matte/Bead-blasted",
        "tags": "Structural, Lightweight-feel, Engineered, Modern",
        "products": "Sporting Goods, Automotive / Transport, Consumer Electronics",
        "applications": "High-Performance / Engineered, Luxury / Premium",
        "rationale": "Exposes the technical weave pattern without reflective glare. UV-stable matte polyurethane clear coat protects the epoxy matrix.",
        "sources": "Ashby & Johnson 2014 [3], CES EduPack [4], Custompart.net [5]",
        "processes": "Vacuum Bag Consolidation; Clear Coat Spraying (PU/Acrylic)"
    },
    {
        "id": 8, "family": "Metals",
        "name": "Brass", "grade": "C36000 (Free-Cutting Brass)",
        "surface": "Chemical Patina (Liver of Sulphur)",
        "Ra": "1.2–3.5 µm", "GU": "4–15 GU",
        "finish": "Industrial/As-built",
        "tags": "Heritage, Warm, Aged, Crafted",
        "products": "Kitchenware / Cutlery, Architectural Hardware, Furniture",
        "applications": "Luxury / Premium, Casual / Home Use, Sustainable / Eco-conscious",
        "rationale": "Controlled oxidation produces mottled brown-to-black tones on a warm yellow base. Wax sealant stabilises the patina.",
        "sources": "ASM Handbook Vol.5 [1], Ashby & Johnson 2014 [3], Karana et al. 2009 [6]",
        "processes": "Chemical Patination; Lacquer / Wax Seal"
    },
    {
        "id": 9, "family": "Polymers",
        "name": "ABS / PC Blend", "grade": "Flame Retardant (UL94 V-0)",
        "surface": "EDM Textured Mold — VDI 24 Fine Texture",
        "Ra": "1.6–2.5 µm", "GU": "3–8 GU",
        "finish": "Matte/Bead-blasted",
        "tags": "Robust, Soft-touch, Functional, Matte",
        "products": "Consumer Electronics, Industrial Tools, Medical Devices, Automotive / Transport",
        "applications": "High-Performance / Engineered, Office / Workspace, Casual / Home Use",
        "rationale": "Controlled spark-eroded mold cavity creates a non-directional micro-texture that is non-slip and hides mold parting lines.",
        "sources": "CES EduPack [4], ASM Handbook Vol.5 [1], Custompart.net [5]",
        "processes": "Injection Molding (VDI/MT Texture); Pad Printing / Laser Marking"
    },
    {
        "id": 10, "family": "Natural Materials",
        "name": "Moso Bamboo", "grade": "Vertical Grain Laminate — Carbonised",
        "surface": "Fine Sanded + Water-based PU Matte",
        "Ra": "1.0–2.0 µm", "GU": "8–18 GU",
        "finish": "Matte/Bead-blasted",
        "tags": "Sustainable, Linear, Zen, Eco-modern",
        "products": "Furniture, Kitchenware / Cutlery, Consumer Electronics, Toys / Recreational",
        "applications": "Sustainable / Eco-conscious, Casual / Home Use, Office / Workspace",
        "rationale": "Matures in 3–5 years vs. 30+ for hardwoods. Vertical lamination emphasises linear grain; carbonisation deepens colour to caramel tones.",
        "sources": "Ashby & Johnson 2014 [3], CES EduPack [4], Karana et al. 2009 [6]",
        "processes": "Precision Sanding (Multi-grit); Clear Coat Spraying (PU/Acrylic)"
    },
    {
        "id": 11, "family": "Metals",
        "name": "Tool Steel", "grade": "AISI H13 (1.2344)",
        "surface": "Shot Peened — Almen A 0.010–0.014″",
        "Ra": "2.0–6.3 µm", "GU": "2–8 GU",
        "finish": "Industrial/As-built",
        "tags": "Rugged, Structural, Industrial, Tough",
        "products": "Industrial Tools, Automotive / Transport",
        "applications": "High-Performance / Engineered, Outdoor / Harsh Environment",
        "rationale": "Introduces compressive residual stress of 400–600 MPa improving fatigue life by 200–600%. Coarse dimpled texture signals heavy-duty durability.",
        "sources": "ASM Handbook Vol.5 [1], ISO 4287 [2], Custompart.net [5]",
        "processes": "Shot Peening; Black Oxide Coating"
    },
    {
        "id": 12, "family": "Polymers",
        "name": "ABS", "grade": "General Purpose / High-Impact",
        "surface": "Bead-Blast Textured Mold — SPI D-2",
        "Ra": "0.8–1.8 µm", "GU": "8–22 GU",
        "finish": "Matte/Bead-blasted",
        "tags": "Neutral, Accessible, Lightweight-feel, Clean",
        "products": "Consumer Electronics, Toys / Recreational, Sporting Goods",
        "applications": "Casual / Home Use, Office / Workspace",
        "rationale": "Cost-effective with excellent impact resistance. Bead-blasted mold texture provides a premium 'soft' look to commodity plastic.",
        "sources": "CES EduPack [4], Ashby & Johnson 2014 [3], Custompart.net [5]",
        "processes": "Injection Molding (VDI/MT Texture); Spray Painting"
    },
    {
        "id": 13, "family": "Composites",
        "name": "Glass Fiber (GFRP)", "grade": "E-Glass / Polyester Hand Layup",
        "surface": "Gel Coat — Semi-gloss White",
        "Ra": "0.3–0.8 µm", "GU": "55–75 GU",
        "finish": "Satin/Semi-gloss",
        "tags": "Durable, Weatherproof, Clean, Lightweight-feel",
        "products": "Automotive / Transport, Sporting Goods, Architectural Hardware, Furniture, Toys / Recreational",
        "applications": "Outdoor / Harsh Environment, High-Performance / Engineered, Sustainable / Eco-conscious",
        "rationale": "Gel coat provides a smooth, pigmented surface integrally bonded to the laminate. Excellent UV and weather resistance.",
        "sources": "Ashby & Johnson 2014 [3], CES EduPack [4], Custompart.net [5]",
        "processes": "Gel Coat Application; Wet Sanding + Buffing"
    },
    {
        "id": 14, "family": "Metals",
        "name": "Copper", "grade": "C11000 (ETP — Electrolytic Tough Pitch)",
        "surface": "Bright Acid Dip + Clear Lacquer",
        "Ra": "0.05–0.2 µm", "GU": "70–90 GU",
        "finish": "Mirror/High-gloss",
        "tags": "Warm, Premium, Antimicrobial, Heritage",
        "products": "Architectural Hardware, Kitchenware / Cutlery, Wearables / Watches",
        "applications": "Luxury / Premium, Casual / Home Use, Sustainable / Eco-conscious",
        "rationale": "Inherent antimicrobial property (EPA registered) eliminates 99.9% of bacteria within 2 hours. Bright dip reveals copper's characteristic warm pink-orange tone.",
        "sources": "ASM Handbook Vol.5 [1], CES EduPack [4], Ashby & Johnson 2014 [3]",
        "processes": "Bright Acid Dip; Clear Lacquer Coat"
    },
    {
        "id": 15, "family": "Metals",
        "name": "Mild Steel", "grade": "ASTM A36 / EN S235",
        "surface": "Sandblasted + Powder Coated (RAL)",
        "Ra": "0.8–2.5 µm", "GU": "15–85 GU",
        "finish": "Satin/Semi-gloss",
        "tags": "Versatile, Durable, Colourful, Accessible",
        "products": "Furniture, Architectural Hardware, Industrial Tools, Toys / Recreational, Kitchenware / Cutlery",
        "applications": "Casual / Home Use, Outdoor / Harsh Environment, Office / Workspace",
        "rationale": "Powder coating creates a thick (60–120 µm), chip-resistant layer available in virtually any RAL colour. Excellent corrosion protection at low cost.",
        "sources": "ASM Handbook Vol.5 [1], Ashby & Johnson 2014 [3], Custompart.net [5]",
        "processes": "Abrasive Blasting (Pre-treatment); Electrostatic Powder Coating"
    },
    {
        "id": 16, "family": "Polymers",
        "name": "Nylon (PA 12)", "grade": "SLS-Grade — 50 µm layer",
        "surface": "As-printed + Vapor Smoothed",
        "Ra": "1.5–4.0 µm", "GU": "5–30 GU ⚠",
        "finish": "Matte/Bead-blasted",
        "tags": "Engineered, Functional, Modern, Technical",
        "products": "Consumer Electronics, Medical Devices, Sporting Goods, Industrial Tools",
        "applications": "High-Performance / Engineered, Office / Workspace",
        "rationale": "SLS produces complex geometries impossible with injection molding. Vapor smoothing reduces texture from Ra ~15 µm to 1.5–4 µm while sealing porosity.",
        "sources": "CES EduPack [4], Custompart.net [5]",
        "processes": "SLS 3D Printing; Chemical Vapor Smoothing"
    },
    {
        "id": 17, "family": "Natural Materials",
        "name": "Teak (Tectona grandis)", "grade": "Plantation-Grown / FSC Certified",
        "surface": "Marine Sanded + Teak Oil Saturated",
        "Ra": "2.0–4.5 µm", "GU": "4–12 GU",
        "finish": "Industrial/As-built",
        "tags": "Weatherproof, Warm, Heritage, Sustainable, Organic",
        "products": "Furniture, Sporting Goods, Architectural Hardware",
        "applications": "Outdoor / Harsh Environment, Luxury / Premium, Sustainable / Eco-conscious",
        "rationale": "Naturally high silica and oil content provides exceptional rot and UV resistance without chemical treatment. Golden-brown tone weathers to silver-grey if unsealed.",
        "sources": "Ashby & Johnson 2014 [3], CES EduPack [4], Karana et al. 2009 [6]",
        "processes": "Precision Sanding (Multi-grit); Natural Oil Saturation (Teak/Tung)"
    },
    {
        "id": 18, "family": "Polymers",
        "name": "PMMA (Acrylic — Plexiglas®)", "grade": "Cast Optical Grade — Cell Cast",
        "surface": "Flame Polished Edges + As-cast Faces",
        "Ra": "0.01–0.03 µm", "GU": "95–100 GU",
        "finish": "Mirror/High-gloss",
        "tags": "Transparent, Premium, Pure, Futuristic, Lightweight-feel",
        "products": "Architectural Hardware, Furniture, Consumer Electronics, Toys / Recreational",
        "applications": "Luxury / Premium, Casual / Home Use, Office / Workspace",
        "rationale": "92% light transmission — optically superior to glass. Cell-cast PMMA has 10× the impact resistance of glass and can be flame-polished to optically clear edges.",
        "sources": "CES EduPack [4], Ashby & Johnson 2014 [3], Custompart.net [5]",
        "processes": "Flame Polishing; CNC Diamond-Fly Cutting"
    },
    {
        "id": 19, "family": "Polymers",
        "name": "Polypropylene (PP)", "grade": "Copolymer — FDA / EU 10/2011 Food Contact",
        "surface": "SPI B-1 Semi-gloss Mold Finish",
        "Ra": "0.05–0.5 µm", "GU": "50–75 GU",
        "finish": "Satin/Semi-gloss",
        "tags": "Clean, Accessible, Lightweight-feel, Functional",
        "products": "Kitchenware / Cutlery, Medical Devices, Toys / Recreational, Consumer Electronics",
        "applications": "Medical / Hygienic, Casual / Home Use, Sustainable / Eco-conscious",
        "rationale": "Excellent chemical resistance with living-hinge capability. Semi-gloss mold finish provides a hygienic, easy-clean surface meeting food-contact regulations.",
        "sources": "CES EduPack [4], Ashby & Johnson 2014 [3], Custompart.net [5]",
        "processes": "Injection Molding (VDI/MT Texture); In-Mold Labeling (IML)"
    },
    {
        "id": 20, "family": "Polymers",
        "name": "Liquid Silicone Rubber (LSR)", "grade": "Medical / Food Grade (Platinum Cured)",
        "surface": "LSR Molded — SPI A-3 Semi-gloss",
        "Ra": "0.05–0.3 µm", "GU": "40–70 GU",
        "finish": "Satin/Semi-gloss",
        "tags": "Soft-touch, Clean, Functional, Durable, Antimicrobial",
        "products": "Wearables / Watches, Medical Devices, Kitchenware / Cutlery, Sporting Goods",
        "applications": "Medical / Hygienic, Casual / Home Use, High-Performance / Engineered, Outdoor / Harsh Environment",
        "rationale": "Biocompatible, hypoallergenic, and temperature-stable from −60°C to +230°C. Gentle elastomeric feel ideal for wearable bands and medical interfaces.",
        "sources": "CES EduPack [4], Ashby & Johnson 2014 [3], Custompart.net [5]",
        "processes": "Liquid Injection Molding (LIM); Overmolding (2K / Insert)"
    },
    {
        "id": 21, "family": "Metals",
        "name": "Aluminum Alloy (High-Strength)", "grade": "AA 7075-T6",
        "surface": "Brushed Satin + Hard Anodize (Type III)",
        "Ra": "0.3–0.8 µm", "GU": "15–35 GU",
        "finish": "Brushed/Directional",
        "tags": "Engineered, Tough, Directional, High-tech, Durable",
        "products": "Sporting Goods, Automotive / Transport, Wearables / Watches, Industrial Tools",
        "applications": "High-Performance / Engineered, Outdoor / Harsh Environment, Luxury / Premium",
        "rationale": "Tensile strength of 572 MPa — comparable to many steels at one-third the density. Brushed satin + Type III hard anodize creates extremely wear-resistant surface.",
        "sources": "ASM Handbook Vol.5 [1], ISO 4287 [2], CES EduPack [4]",
        "processes": "Abrasive Belt Brushing; Type III Hard Anodize"
    },
]

# ── Process database ──
PROCESSES = [
    {"name": "Bead Blasting",                     "Ra": "0.8–2.0 µm", "GU": "8–25 GU",   "aesthetic": "Satin-matte, isotropic scattering",        "compatibility": "Al, SS, Ti alloys. Avoid thin-walled parts < 0.5 mm.", "setup": "Low",    "unit_cost": "$2–8/unit",              "tier": "★★☆☆☆", "source": "[1, 5]"},
    {"name": "Type II Anodizing",                  "Ra": "Substrate ± 0.1 µm", "GU": "Substrate dep.", "aesthetic": "Metallic depth, colorable; excellent UV stability", "compatibility": "Aluminum alloys only. Not for high-Si castings (> 7% Si).", "setup": "Medium", "unit_cost": "$3–10/unit",             "tier": "★★★☆☆", "source": "[1, 5]"},
    {"name": "Type III Hard Anodize",              "Ra": "Substrate + 0.2 µm", "GU": "5–15 GU",  "aesthetic": "Dark grey/black, industrial, extremely hard", "compatibility": "Aluminum alloys only. Not rec. for 2xxx/7xxx with > 5% Cu.", "setup": "High",   "unit_cost": "$8–20/unit",             "tier": "★★★★☆", "source": "[1, 5]"},
    {"name": "Electropolishing",                   "Ra": "< 0.1 µm",   "GU": "> 90 GU",  "aesthetic": "Brilliant, sterile, ultra-smooth isotropic",  "compatibility": "Best for austenitic SS (304, 316). Difficult on ferritic.", "setup": "Medium", "unit_cost": "$5–15/unit",             "tier": "★★★★☆", "source": "[1, 5]"},
    {"name": "Fine Lapping",                       "Ra": "0.02–0.05 µm", "GU": "90–100 GU", "aesthetic": "Optical-grade mirror; zero visible texture",   "compatibility": "Any metal. Very slow; limited to flat/convex surfaces.",  "setup": "High",   "unit_cost": "$20–60/unit",            "tier": "★★★★★", "source": "[1, 5]"},
    {"name": "Abrasive Belt Brushing",             "Ra": "0.4–1.0 µm", "GU": "20–50 GU",  "aesthetic": "Directional 'grain' lines; conceals wear",   "compatibility": "Most metals. Difficult on complex 3D curves.",           "setup": "Low",    "unit_cost": "$1–4/unit",              "tier": "★☆☆☆☆", "source": "[1, 5]"},
    {"name": "Passivation (Citric Acid)",          "Ra": "No change",   "GU": "No change", "aesthetic": "Invisible; restores passive Cr₂O₃ layer",    "compatibility": "Stainless steel only. Required after machining/welding.", "setup": "Low",    "unit_cost": "$1–3/unit",              "tier": "★☆☆☆☆", "source": "[1, 5]"},
    {"name": "PVD Coating (TiN/TiAlN/CrN)",       "Ra": "Substrate ± 0.05 µm", "GU": "60–90 GU", "aesthetic": "Thin-film interference colors (gold/blue/silver)", "compatibility": "Any metal or ceramic. Requires vacuum chamber.", "setup": "High",   "unit_cost": "$15–50/unit",            "tier": "★★★★★", "source": "[1, 5]"},
    {"name": "Precision Grinding",                 "Ra": "0.1–0.4 µm", "GU": "50–80 GU",  "aesthetic": "Flat, reflective; controlled geometry",       "compatibility": "All metals. Limited to accessible flat/cylindrical.",     "setup": "Medium", "unit_cost": "$5–15/unit",             "tier": "★★★☆☆", "source": "[1, 5]"},
    {"name": "Precision Injection Molding (SPI A-1)","Ra": "0.012–0.05 µm", "GU": "90–100 GU", "aesthetic": "Liquid-smooth, mirror-like from diamond mold", "compatibility": "PC, PMMA, ABS. Requires SPI A-1/A-2 mold finish.", "setup": "High",   "unit_cost": "$0.50–2/unit (amort.)",  "tier": "★★☆☆☆", "source": "[4, 5]"},
    {"name": "Vapor Polishing (Acetone/DCM)",      "Ra": "< 0.05 µm",  "GU": "95–100 GU", "aesthetic": "Chemical smoothing; glass-like clarity",       "compatibility": "PC, ABS, PMMA. Not for PE, PP. Requires fume extraction.","setup": "Low",    "unit_cost": "$1–5/unit",              "tier": "★★☆☆☆", "source": "[4, 5]"},
    {"name": "Vacuum Bag Consolidation",           "Ra": "0.3–1.0 µm (tool side)", "GU": "5–60 GU", "aesthetic": "Controlled fiber compaction; visible weave", "compatibility": "All fiber-reinforced thermoset composites.", "setup": "Medium", "unit_cost": "$10–30/unit",            "tier": "★★★★☆", "source": "[3, 5]"},
    {"name": "Clear Coat Spraying (PU/Acrylic)",   "Ra": "0.2–1.5 µm", "GU": "5–90 GU",   "aesthetic": "Depth and protection; matte to gloss options", "compatibility": "All composites and wood surfaces. Dust-free env. required.","setup": "Low",    "unit_cost": "$3–8/unit",              "tier": "★★☆☆☆", "source": "[5]"},
    {"name": "Chemical Patination",                "Ra": "1.0–4.0 µm", "GU": "< 15 GU",   "aesthetic": "Mottled, organic oxidation; unique each piece", "compatibility": "Copper-based alloys only (brass, bronze).", "setup": "Low",    "unit_cost": "$4–12/unit",             "tier": "★★☆☆☆", "source": "[1, 3]"},
    {"name": "Lacquer / Wax Seal",                 "Ra": "Slight smoothing", "GU": "+5–15 GU", "aesthetic": "Locks patina; adds warm depth; prevents oxidation", "compatibility": "All metals and wood. Re-apply every 6–12 months.", "setup": "Low",    "unit_cost": "$2–5/unit",              "tier": "★☆☆☆☆", "source": "[3, 5]"},
    {"name": "Injection Molding (VDI/MT Texture)", "Ra": "0.4–12 µm",  "GU": "2–60 GU",   "aesthetic": "Isotropic matte to semi-gloss from EDM'd mold", "compatibility": "ABS, PC, PP, PA. Min draft 1°/0.025mm texture depth.", "setup": "High",   "unit_cost": "$0.30–1.50/unit (amort.)","tier": "★★☆☆☆", "source": "[4, 5]"},
    {"name": "Shot Peening",                       "Ra": "1.6–6.3 µm", "GU": "< 10 GU",   "aesthetic": "Dimpled, non-reflective, visibly rugged",     "compatibility": "Most structural metals. Avoid on thin sections < 1 mm.", "setup": "Medium", "unit_cost": "$5–12/unit",             "tier": "★★★☆☆", "source": "[1, 5]"},
    {"name": "Black Oxide Coating",                "Ra": "No significant change", "GU": "< 5 GU", "aesthetic": "Uniform matte black; zero dimensional change", "compatibility": "Carbon and alloy steels only. Not for SS or Al.", "setup": "Low",    "unit_cost": "$2–6/unit",              "tier": "★★☆☆☆", "source": "[1, 5]"},
    {"name": "Spray Painting",                     "Ra": "0.5–2.0 µm", "GU": "5–95 GU",   "aesthetic": "Unlimited color; soft-touch, metallic, pearlescent", "compatibility": "All polymers (with primer). Poor on PE/PP without flame treat.", "setup": "Medium", "unit_cost": "$2–8/unit",              "tier": "★★★☆☆", "source": "[5]"},
    {"name": "Gel Coat Application",               "Ra": "0.2–0.8 µm", "GU": "50–80 GU",  "aesthetic": "Smooth, opaque, automotive-quality from mold", "compatibility": "Polyester and vinyl ester laminates. Not epoxy without primer.","setup": "Low",    "unit_cost": "$3–10/unit",             "tier": "★★☆☆☆", "source": "[3, 5]"},
    {"name": "Wet Sanding + Buffing",              "Ra": "< 0.1 µm",   "GU": "80–95 GU",  "aesthetic": "Show-car finish; removes orange peel",        "compatibility": "Gel-coated parts only. Labour-intensive; not scalable.",  "setup": "Low",    "unit_cost": "$8–20/unit",             "tier": "★★★☆☆", "source": "[5]"},
    {"name": "Bright Acid Dip",                    "Ra": "0.05–0.2 µm","GU": "70–95 GU",  "aesthetic": "Chemical polish; uniform warmth; no directional marks", "compatibility": "Copper and high-Cu alloys only. Aggressive on Zn-rich brass.","setup": "Low",    "unit_cost": "$3–8/unit",              "tier": "★★☆☆☆", "source": "[1, 5]"},
    {"name": "Electrostatic Powder Coating",       "Ra": "0.8–3.0 µm", "GU": "15–95 GU",  "aesthetic": "Uniform colour; matte/satin/gloss; textured options", "compatibility": "Electrically conductive substrates. Max ~200°C service.", "setup": "Medium", "unit_cost": "$3–10/unit",             "tier": "★★☆☆☆", "source": "[1, 5]"},
    {"name": "Abrasive Blasting (Pre-treatment)",  "Ra": "2.0–6.0 µm", "GU": "< 10 GU",   "aesthetic": "Anchor profile for coating adhesion",          "compatibility": "All ferrous metals. Profile matched to coating thickness.","setup": "Low",    "unit_cost": "$2–6/unit",              "tier": "★★☆☆☆", "source": "[1, 5]"},
    {"name": "SLS 3D Printing",                    "Ra": "6–16 µm",    "GU": "< 5 GU",    "aesthetic": "Granular, powder-bed texture; visible layer lines", "compatibility": "PA12, PA11, TPU. Not for transparent parts.", "setup": "High",   "unit_cost": "$8–40/unit",             "tier": "★★★★☆", "source": "[4, 5]"},
    {"name": "Chemical Vapor Smoothing",           "Ra": "1.5–4.0 µm", "GU": "10–35 GU ⚠","aesthetic": "Semi-gloss sealed; approaches injection-molded look", "compatibility": "PA12, PA11. Proprietary chemistry (AMT PostPro).", "setup": "Medium", "unit_cost": "$5–15/unit",             "tier": "★★★☆☆", "source": "[4, 5]"},
    {"name": "Precision Sanding (Multi-grit)",     "Ra": "1.6–6.3 µm", "GU": "5–20 GU",   "aesthetic": "Reveals grain depth; silky tactile feel",      "compatibility": "All hardwoods and softwoods. Grain raising on ring-porous.","setup": "Low",    "unit_cost": "$2–6/unit",              "tier": "★☆☆☆☆", "source": "[3, 5]"},
    {"name": "Hand Oil / Wax Application",         "Ra": "Slight reduction", "GU": "8–25 GU", "aesthetic": "Natural luster; enhances grain contrast and color", "compatibility": "All wood species. Drying 24–72 hrs. Food-safe if specified.","setup": "Low",    "unit_cost": "$3–8/unit",              "tier": "★★☆☆☆", "source": "[3, 5]"},
    {"name": "Natural Oil Saturation (Teak/Tung)", "Ra": "Slight reduction", "GU": "5–15 GU", "aesthetic": "Deepens color; water-beading surface",         "compatibility": "All tropical/temperate hardwoods. Re-oil annually outdoors.","setup": "Low",    "unit_cost": "$3–8/unit",              "tier": "★★☆☆☆", "source": "[3, 5]"},
    {"name": "Flame Polishing",                    "Ra": "< 0.02 µm",  "GU": "95–100 GU", "aesthetic": "Crystal-clear edge; melts micro-roughness",    "compatibility": "PMMA and PC only. Thickness > 3 mm. Fire risk.",         "setup": "Low",    "unit_cost": "$1–4/unit",              "tier": "★☆☆☆☆", "source": "[4, 5]"},
    {"name": "CNC Diamond-Fly Cutting",            "Ra": "< 0.01 µm",  "GU": "98–100 GU", "aesthetic": "Optical-grade flat mirror directly from tool",  "compatibility": "PMMA, PC, soft metals (Al, Cu). Diamond tooling very expensive.","setup": "High",   "unit_cost": "$10–40/unit",            "tier": "★★★★★", "source": "[1, 5]"},
    {"name": "In-Mold Labeling (IML)",             "Ra": "Label dependent", "GU": "Label dep.", "aesthetic": "Photorealistic graphics fused into surface", "compatibility": "PP and PE primarily. Label must match resin for recycling.","setup": "Medium", "unit_cost": "$0.05–0.30/label",       "tier": "★☆☆☆☆", "source": "[5]"},
    {"name": "Liquid Injection Molding (LIM)",      "Ra": "0.05–0.5 µm","GU": "30–80 GU",  "aesthetic": "Translucent/opaque; skin-like; parting line invisible", "compatibility": "LSR and HCR silicones only. Specialized LIM press required.","setup": "High",   "unit_cost": "$0.50–3/unit (amort.)",  "tier": "★★★☆☆", "source": "[4, 5]"},
    {"name": "Overmolding (2K / Insert)",           "Ra": "Substrate + LSR", "GU": "Mixed",  "aesthetic": "Rigid + soft-touch grip zones; dual-material aesthetic", "compatibility": "LSR bonds to PC, PA, PBT with primer. Poor on PP, PE.", "setup": "High",   "unit_cost": "$1–5/unit",              "tier": "★★★☆☆", "source": "[4, 5]"},
    {"name": "Pad Printing / Laser Marking",       "Ra": "Localised",   "GU": "Localised", "aesthetic": "Branding, icons, regulatory marks on textured surface", "compatibility": "All polymers. Laser marking needs contrast testing per material.","setup": "Medium", "unit_cost": "$0.50–3/unit",           "tier": "★★☆☆☆", "source": "[5]"},
]


# ── Styling helpers ──
thin_border = Border(
    left=Side(style='thin', color=BORDER_CLR),
    right=Side(style='thin', color=BORDER_CLR),
    top=Side(style='thin', color=BORDER_CLR),
    bottom=Side(style='thin', color=BORDER_CLR),
)

header_font = Font(name='Calibri', size=11, bold=True, color=WHITE)
header_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
subheader_font = Font(name='Calibri', size=10, bold=True)
body_font = Font(name='Calibri', size=10)
wrap_align = Alignment(wrap_text=True, vertical='top')
center_align = Alignment(horizontal='center', vertical='top', wrap_text=True)

MAT_HEADERS = [
    "ID", "Material Name", "Grade / Alloy", "Surface Condition",
    "Ra Range", "Gloss (60°)", "Finish Category",
    "Aesthetic Tags", "Product Lines", "Applications",
    "Rationale", "Processes", "Sources"
]
MAT_WIDTHS = [5, 22, 26, 34, 16, 14, 20, 32, 44, 40, 55, 38, 38]

PROC_HEADERS = [
    "Process Name", "Achievable Ra", "Achievable Gloss (60°)",
    "Aesthetic Character", "Material Compatibility",
    "Setup Cost", "Unit Cost Range", "Cost Tier", "Source"
]
PROC_WIDTHS = [34, 22, 22, 42, 48, 12, 22, 14, 12]


def style_header_row(ws, headers, widths, row=1):
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[row].height = 30


def write_mat_row(ws, row, m, fill=None):
    vals = [
        m["id"], m["name"], m["grade"], m["surface"],
        m["Ra"], m["GU"], m["finish"],
        m["tags"], m["products"], m["applications"],
        m["rationale"], m["processes"], m["sources"]
    ]
    for col, v in enumerate(vals, 1):
        cell = ws.cell(row=row, column=col, value=v)
        cell.font = body_font
        cell.alignment = wrap_align
        cell.border = thin_border
        if fill:
            cell.fill = fill
    ws.row_dimensions[row].height = 60


def write_proc_row(ws, row, p, fill=None):
    vals = [
        p["name"], p["Ra"], p["GU"],
        p["aesthetic"], p["compatibility"],
        p["setup"], p["unit_cost"], p["tier"], p["source"]
    ]
    for col, v in enumerate(vals, 1):
        cell = ws.cell(row=row, column=col, value=v)
        cell.font = body_font
        cell.alignment = wrap_align
        cell.border = thin_border
        if fill:
            cell.fill = fill
    ws.row_dimensions[row].height = 48


def create_workbook():
    wb = openpyxl.Workbook()

    # ── Tab 1: All Materials ──
    ws_all = wb.active
    ws_all.title = "All Materials (21)"
    ws_all.sheet_properties.tabColor = BLUE
    style_header_row(ws_all, MAT_HEADERS, MAT_WIDTHS)
    stripe = PatternFill(start_color=LIGHT_GREY, end_color=LIGHT_GREY, fill_type='solid')
    for i, m in enumerate(MATERIALS):
        write_mat_row(ws_all, i + 2, m, fill=(stripe if i % 2 == 1 else None))
    ws_all.freeze_panes = 'A2'
    ws_all.auto_filter.ref = f"A1:M{len(MATERIALS)+1}"

    # ── Tabs 2–5: By Family ──
    families_order = ["Metals", "Polymers", "Natural Materials", "Composites"]
    for fam in families_order:
        fam_mats = [m for m in MATERIALS if m["family"] == fam]
        tab_name = f"{fam} ({len(fam_mats)})"
        ws = wb.create_sheet(title=tab_name)
        ws.sheet_properties.tabColor = FAMILY_HEADER_COLORS[fam]

        fam_fill = PatternFill(start_color=FAMILY_COLORS[fam], end_color=FAMILY_COLORS[fam], fill_type='solid')

        style_header_row(ws, MAT_HEADERS, MAT_WIDTHS)
        for i, m in enumerate(fam_mats):
            write_mat_row(ws, i + 2, m, fill=(fam_fill if i % 2 == 0 else None))
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f"A1:M{len(fam_mats)+1}"

    # ── Tab 6: All Processes ──
    ws_proc = wb.create_sheet(title=f"Processes ({len(PROCESSES)})")
    ws_proc.sheet_properties.tabColor = "10B981"
    style_header_row(ws_proc, PROC_HEADERS, PROC_WIDTHS)
    for i, p in enumerate(PROCESSES):
        write_proc_row(ws_proc, i + 2, p, fill=(stripe if i % 2 == 1 else None))
    ws_proc.freeze_panes = 'A2'
    ws_proc.auto_filter.ref = f"A1:I{len(PROCESSES)+1}"

    return wb


if __name__ == "__main__":
    wb = create_workbook()
    path = "/Users/priyanshgupta/Desktop/MITTAL_BTP/Material_Database_CP302.xlsx"
    wb.save(path)
    print(f"✅ Workbook saved to: {path}")
    print(f"   Tabs: {[s.title for s in wb.worksheets]}")
    print(f"   Materials: {len(MATERIALS)} | Processes: {len(PROCESSES)}")
