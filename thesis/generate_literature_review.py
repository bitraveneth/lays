"""Generate the MS-thesis literature review chapter as an APA-7 formatted .docx.

Every reference in the bibliography was verified against publisher/indexing
records (DOI, journal, volume, issue, pages) before inclusion.
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Times New Roman"

doc = Document()

# ---------------------------------------------------------------- base style
style = doc.styles["Normal"]
style.font.name = FONT
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
pf = style.paragraph_format
pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
pf.space_after = Pt(0)
pf.space_before = Pt(0)

for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

# ------------------------------------------------------- page number header
header_p = doc.sections[0].header.paragraphs[0]
header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = header_p.add_run()
fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_end)
run.font.name = FONT
run.font.size = Pt(12)


# ------------------------------------------------------------------ helpers
def body(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def h1(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = True
    return p


def h2(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    return p


def h3(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.italic = True
    return p


def ref(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    return p


def page_break():
    doc.add_page_break()


# ================================================================ TITLE PAGE
for _ in range(4):
    doc.add_paragraph()
tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("Modeling Rainfall\u2013Runoff in the Gumai Beel, Ichamoti River Using HEC-HMS")
r.bold = True
for line in [
    "",
    "Chapter 2: Literature Review",
    "",
    "[Student Name]",
    "Department of [Department Name], [University Name]",
    "MS Thesis",
    "[Supervisor Name], Supervisor",
    "[Month, Year]",
]:
    p = doc.add_paragraph(line)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
page_break()

# ================================================================= CHAPTER 2
h1("Chapter 2: Literature Review")

h2("2.1 Introduction")
body(
    "This chapter reviews the body of scientific literature that underpins the present study, which develops a "
    "rainfall\u2013runoff model of the Gumai Beel catchment within the Ichamoti River system of Pabna District, "
    "Bangladesh, using the Hydrologic Engineering Center\u2019s Hydrologic Modeling System (HEC-HMS). The review is "
    "organized to move from the general to the particular. It begins with the conceptual foundations of rainfall\u2013runoff "
    "processes and the classification of hydrological models, and then examines the structure of the HEC-HMS framework and "
    "the scientific basis of its principal computational methods. Subsequent sections review the role of geographic "
    "information systems (GIS) and remote sensing in supplying model inputs, survey applications of HEC-HMS across diverse "
    "hydro-climatic regions, and narrow the focus to hydrological modeling experience in Bangladesh. Because the study area "
    "is not an ordinary upland catchment but a low-lying floodplain wetland (beel) drained by a moribund distributary of the "
    "Padma River, a dedicated section reviews the hydrology of beel and floodplain-wetland systems, the anthropogenic decline "
    "of the Ichamoti River, and the chronic drainage congestion of the Pabna region. The chapter then reviews accepted "
    "standards for model calibration, validation, and performance evaluation, and closes with a synthesis that identifies the "
    "research gap this thesis addresses."
)
body(
    "Throughout the chapter, emphasis is placed on three questions that guide the design of the present study: (a) which "
    "combinations of loss, transform, baseflow, and routing methods within HEC-HMS have proven reliable in monsoon-dominated "
    "and data-scarce catchments; (b) what data sources and pre-processing strategies are appropriate where gauge networks are "
    "sparse and the terrain is extremely flat; and (c) what performance criteria constitute an acceptable model in the context "
    "of an applied water-management problem such as the drainage rehabilitation of the Gumai Beel\u2013Ichamoti system."
)

h2("2.2 Rainfall\u2013Runoff Processes and the Classification of Hydrological Models")
body(
    "The transformation of rainfall into streamflow is the central problem of applied hydrology. Precipitation falling on a "
    "catchment is partitioned among interception, depression storage, infiltration, evapotranspiration, soil-moisture "
    "replenishment, and surface runoff, and the runoff component is subsequently translated and attenuated as it travels "
    "over hillslopes and through channel networks to the catchment outlet (Chow et al., 1988). Because direct measurement of "
    "each of these processes at the catchment scale is impossible, hydrologists rely on mathematical models that represent "
    "the catchment as a system converting an input hyetograph into an output hydrograph. Singh and Woolhiser (2002), in a "
    "comprehensive review of watershed modeling, traced the evolution of such models from the rational method and unit "
    "hydrograph theory of the early twentieth century to the distributed, physics-based codes of the modern era, and observed "
    "that model complexity must always be balanced against data availability and the purpose of the simulation."
)
body(
    "Hydrological models are conventionally classified along two axes. The first concerns the treatment of process "
    "representation: empirical (or metric) models relate input to output through statistically fitted relationships with no "
    "claim to physical realism; conceptual models represent the catchment as a series of interconnected conceptual storages "
    "whose behavior mimics physical processes; and physically based models solve equations of mass, momentum, and energy "
    "conservation with parameters that are, in principle, measurable (Devia et al., 2015). The second axis concerns spatial "
    "discretization: lumped models treat the catchment as a single homogeneous unit, semi-distributed models divide it into "
    "sub-basins within which parameters are lumped, and fully distributed models resolve the catchment on a regular grid "
    "(Beven, 2012). No single position in this classification space is universally superior. Physically based distributed "
    "models offer the richest process description but demand data that rarely exist in developing-country catchments, and "
    "their many degrees of freedom can produce an illusion of realism that the available observations cannot constrain "
    "(Beven, 2012). Conceptual semi-distributed models, by contrast, require modest data, are computationally inexpensive, "
    "and have repeatedly been shown to reproduce observed hydrographs with accuracy sufficient for planning and design "
    "purposes (Devia et al., 2015; Singh & Woolhiser, 2002)."
)
body(
    "A further operational distinction separates event-based from continuous simulation. Event-based modeling simulates the "
    "catchment response to an individual storm and is appropriate for design-flood estimation and flood forecasting, whereas "
    "continuous modeling tracks moisture accounting through wet and dry periods over months or years and is appropriate for "
    "water-balance and water-resources assessment (Chu & Steinman, 2009). Chu and Steinman (2009) demonstrated, for the Mona "
    "Lake watershed in Michigan, that the two modes are complementary: parameters identified through fine-time-step event "
    "calibration can be transferred to a coarse-time-step continuous model, improving its reliability where long-term "
    "intensive monitoring data are unavailable. This complementarity is directly relevant to the present study, in which "
    "monsoon-season runoff events superimposed on a strongly seasonal water balance govern the inundation behavior of the "
    "Gumai Beel."
)
body(
    "For a data-scarce, monsoon-dominated floodplain catchment such as the study area, the literature therefore points toward "
    "a conceptual, semi-distributed modeling strategy capable of both event and continuous operation. Among the freely "
    "available modeling systems that satisfy these criteria, HEC-HMS has become one of the most widely adopted worldwide, and "
    "the following section reviews its structure and methods in detail."
)

h2("2.3 The HEC-HMS Modeling Framework")

h3("2.3.1 Origin and Structure")
body(
    "HEC-HMS was developed by the Hydrologic Engineering Center of the U.S. Army Corps of Engineers as the successor to the "
    "HEC-1 flood hydrograph package, and it is designed to simulate the complete precipitation\u2013runoff process of dendritic "
    "watershed systems (U.S. Army Corps of Engineers [USACE], 2000). A HEC-HMS project comprises three principal components: "
    "a basin model, which describes the physical catchment as a network of sub-basins, reaches, junctions, reservoirs, "
    "diversions, sources, and sinks; a meteorological model, which assigns precipitation and evapotranspiration boundary "
    "conditions to the sub-basins; and control specifications, which define the simulation window and time step (USACE, 2000). "
    "Within each sub-basin, the user selects one method from each of several interchangeable libraries\u2014canopy and surface "
    "storage, loss (infiltration), direct-runoff transform, and baseflow\u2014while each reach is assigned a channel-routing "
    "method. This modular architecture is the principal reason for the model\u2019s versatility: the same software can be "
    "configured as a simple event-based lumped model or as a gridded, continuous, semi-distributed model. The program is in "
    "the public domain, is supported by extensive documentation, and integrates with GIS-based terrain preprocessing tools "
    "(Fleming & Doan, 2013), all of which explain its popularity in regions where commercial modeling systems are "
    "unaffordable (Oleyiblo & Li, 2010)."
)

h3("2.3.2 Loss Methods and the SCS Curve Number")
body(
    "The loss model determines how much of the incident rainfall infiltrates or is otherwise abstracted, and therefore how "
    "much becomes precipitation excess available for direct runoff. The most widely used loss method in HEC-HMS applications "
    "is the Soil Conservation Service Curve Number (SCS-CN) method, developed by the U.S. Department of Agriculture from "
    "empirical analysis of small agricultural watersheds (USDA Soil Conservation Service, 1985). The method condenses the "
    "combined influence of soil hydrologic group, land use and treatment, surface condition, and antecedent moisture into a "
    "single dimensionless parameter, the curve number (CN), which varies between 0 and 100. Its appeal lies in its "
    "simplicity, its modest data requirements, and its direct linkage to mappable catchment properties, which allows CN "
    "values to be estimated for ungauged areas from land-cover and soil maps (Mishra & Singh, 2003). The method has also "
    "attracted sustained critique. Ponce and Hawkins (1996), in a widely cited appraisal, concluded that while the method is "
    "well established for event-scale runoff estimation on agricultural catchments, its lack of an explicit time dimension "
    "and its sensitivity to antecedent moisture demand caution, particularly outside the range of conditions for which it "
    "was derived. Mishra and Singh (2003) provided a comprehensive theoretical re-derivation of the method and catalogued "
    "numerous modifications intended to extend its validity."
)
body(
    "HEC-HMS offers several alternatives to the SCS-CN approach, including the initial-and-constant loss method, the "
    "deficit-and-constant method, the Green\u2013Ampt infiltration model, and, for continuous simulation, the five-layer soil "
    "moisture accounting (SMA) algorithm (USACE, 2000). The empirical evidence indicates that no loss method is universally "
    "best and that the choice must be tested against local data. Halwatura and Najim (2013), calibrating HEC-HMS 3.4 for the "
    "tropical Attanagalu Oya catchment in Sri Lanka, found that the deficit-and-constant loss method outperformed the SCS-CN "
    "method, which they reported to perform poorly for their wet-zone catchment. Conversely, Gilewski and Nawalany (2018) "
    "found the SCS-CN method to perform best for flood events with unimodal temporal distributions in a mountainous Polish "
    "catchment, and Tassew et al. (2019) obtained very good performance with the SCS-CN method in the monsoonal Gilgel Abay "
    "catchment of Ethiopia. For continuous applications, Chu and Steinman (2009) paired the SCS-CN method for event simulation "
    "with the SMA method for continuous simulation of the same watershed, illustrating how the two can be used in tandem. "
    "These mixed findings justify the comparative testing of loss methods that is undertaken in the methodology of the "
    "present study."
)

h3("2.3.3 Direct-Runoff Transform Methods")
body(
    "The transform model converts precipitation excess into a direct-runoff hydrograph at the sub-basin outlet. Most HEC-HMS "
    "applications employ unit-hydrograph theory, of which three synthetic variants dominate: the SCS dimensionless unit "
    "hydrograph, parameterized by basin lag time; the Clark unit hydrograph, which combines a time\u2013area histogram with a "
    "linear-reservoir storage coefficient to represent translation and attenuation separately; and the Snyder unit "
    "hydrograph, an empirical method parameterized by lag and peaking coefficients (Chow et al., 1988; USACE, 2000). "
    "Comparative evidence again indicates that the appropriate choice is catchment specific. Halwatura and Najim (2013) "
    "reported that the Snyder unit hydrograph simulated flows more reliably than the Clark unit hydrograph in their tropical "
    "catchment, while the SCS unit hydrograph combined with the SCS-CN loss model has been the most frequently adopted "
    "configuration in South Asian and African applications (Nujhat et al., 2024; Tassew et al., 2019). For very flat "
    "catchments, lag-time estimation deserves particular care because empirical lag equations were developed predominantly "
    "for sloping terrain, and misestimated lag propagates directly into errors in peak timing (Chow et al., 1988)."
)

h3("2.3.4 Baseflow Methods")
body(
    "Baseflow representation is frequently the weakest element of event-oriented models, yet in floodplain-wetland systems "
    "the slow drainage component can dominate the recession limb and the dry-season water balance. HEC-HMS provides "
    "recession, bounded-recession, linear-reservoir, and constant-monthly baseflow methods (USACE, 2000). Experience in "
    "Bangladesh suggests that baseflow is a genuine difficulty: M. B. Haque et al. (2024), modeling the Halda River "
    "catchment, obtained satisfactory overall statistics but reported a poor match for the baseflow portion of the "
    "hydrograph during calibration, which they attributed to unrepresented groundwater\u2013surface water exchange, and they "
    "recommended coupling with a groundwater model as a route to improvement. Such findings caution against interpreting "
    "event-calibrated models as complete descriptions of low-flow behavior, a caution of particular force in beel "
    "environments where monsoon storage is released gradually through the post-monsoon season."
)

h3("2.3.5 Channel Routing Methods")
body(
    "Flow routing through reaches is available in HEC-HMS through the Muskingum, Muskingum\u2013Cunge, kinematic-wave, "
    "modified-Puls, and lag methods (USACE, 2000). The Muskingum method, which represents a reach as a linear storage with "
    "travel-time parameter K and weighting parameter X, remains the most commonly adopted in applications comparable to the "
    "present study (Nujhat et al., 2024; Tassew et al., 2019). Its limitation\u2014shared by all hydrologic routing schemes\u2014is "
    "that it cannot represent backwater effects, flow reversal, or looped stage\u2013discharge relations, which are common in "
    "extremely flat deltaic channels. Where such hydraulic effects matter, the literature couples HEC-HMS with a hydraulic "
    "model: Knebl et al. (2005) demonstrated an influential framework in which HEC-HMS-derived hydrographs drive unsteady "
    "HEC-RAS simulation for floodplain mapping of the San Antonio River basin. An analogous coupling represents a natural "
    "extension of the present work, given that drainage of the Gumai Beel is partly controlled by water levels in the "
    "receiving Ichamoti channel rather than by catchment runoff alone."
)

h2("2.4 GIS, Remote Sensing, and Input Data for Hydrological Modeling")

h3("2.4.1 Digital Elevation Models and Watershed Delineation")
body(
    "Semi-distributed modeling begins with terrain analysis: delineation of sub-basins and stream networks from a digital "
    "elevation model (DEM), and extraction of physiographic parameters such as area, slope, and longest flow path. The "
    "Shuttle Radar Topography Mission (SRTM) provided the first near-global, freely available elevation dataset and remains "
    "the most widely used DEM in developing-country hydrology (Farr et al., 2007). The GIS companion tools of HEC-HMS, "
    "notably the HEC-GeoHMS extension and its successors integrated into recent program versions, automate terrain "
    "preprocessing, basin processing, and the assembly of basin-model files (Fleming & Doan, 2013; Oleyiblo & Li, 2010). "
    "The reliability of automated delineation, however, degrades precisely in the terrain type of the present study: "
    "low-relief floodplains. Datta et al. (2022), working on the Halda watershed in Bangladesh, showed systematically that "
    "delineation outcomes depend materially on the choice of DEM product, its spatial resolution, and the stream-definition "
    "area threshold, and that these choices propagate into sub-basin geometry and derived parameters. In flat deltaic "
    "terrain, where total relief may be only a few meters and anthropogenic features such as roads and embankments control "
    "actual flow paths, DEM vertical error can exceed the topographic signal, and the literature therefore recommends "
    "verification of automatically delineated drainage against field knowledge and hydrographic maps (Datta et al., 2022). "
    "This consideration is central to the delineation strategy adopted in Chapter 3 of this thesis."
)

h3("2.4.2 Land Use, Soils, and Curve Number Generation")
body(
    "Loss-model parameterization in ungauged or sparsely gauged basins rests on thematic mapping. The standard workflow "
    "intersects a land-use/land-cover (LULC) classification, commonly derived from Landsat or Sentinel-2 imagery, with a "
    "hydrologic soil group map derived from soil surveys, and assigns composite curve numbers to each sub-basin through "
    "standard lookup tables (Fleming & Doan, 2013; USDA Soil Conservation Service, 1985). Applications across climates "
    "confirm both the practicality and the sensitivity of this workflow: Tassew et al. (2019) identified the curve number "
    "as the single most sensitive parameter of their model, and M. B. Haque et al. (2024) likewise found the "
    "runoff\u2013rainfall relationship to be highly sensitive to sub-basin CN values, which depend directly on the LULC "
    "classification. These findings imply that LULC accuracy assessment and CN calibration bounds deserve explicit "
    "attention in the present study, particularly because beel environments exhibit strong seasonal alternation between "
    "open water, cropped land, and marsh vegetation that a single-date LULC classification cannot capture."
)

h3("2.4.3 Precipitation Inputs: Gauges, Reanalysis, and Satellite Products")
body(
    "Precipitation is the dominant control on simulated runoff, and its estimation is a major source of uncertainty where "
    "gauge networks are sparse. Three families of alternatives to gauge interpolation have matured over the past two "
    "decades: ground-based weather radar, satellite multi-sensor products such as the Integrated Multi-satellitE Retrievals "
    "for Global Precipitation Measurement (IMERG), and atmospheric reanalyses such as ERA5 (Hersbach et al., 2020). The "
    "Climate Hazards Group InfraRed Precipitation with Station data (CHIRPS) offers a long quasi-global record blending "
    "satellite estimates with station observations, designed expressly for data-sparse regions (Funk et al., 2015). "
    "Gilewski and Nawalany (2018) compared rain-gauge, adjusted-radar, and IMERG forcing of an HEC-HMS model and found that "
    "adjusted radar and IMERG were the most reliable sources for event-based modeling, while emphasizing that the model must "
    "be recalibrated separately for each forcing because the spatial and temporal structure of rainfall significantly "
    "affects parameter estimates. In Bangladesh, ERA5 forcing has been used successfully where Bangladesh Meteorological "
    "Department gauge records are sparse or discontinuous (M. B. Haque et al., 2024). For the present study, which can draw "
    "on Bangladesh Water Development Board (BWDB) and Bangladesh Meteorological Department gauges in and around Pabna, the "
    "literature supports a strategy of gauge-primary forcing with satellite or reanalysis products used for gap filling and "
    "consistency checking."
)

h2("2.5 Applications of HEC-HMS Across Diverse Hydro-Climatic Regions")
body(
    "The international literature on HEC-HMS applications is extensive, and a selective review of methodologically "
    "instructive studies is presented here; Table 2.1 summarizes their essential characteristics. In the United States, "
    "Knebl et al. (2005) established the template for regional-scale flood modeling by integrating NEXRAD radar rainfall, "
    "GIS preprocessing, HEC-HMS runoff simulation, and HEC-RAS hydraulic routing across the roughly 10,000 km\u00b2 San "
    "Antonio River basin for the severe summer 2002 storm, demonstrating that manually calibrated sub-basin parameters "
    "could reproduce discharge at twelve interior points and support credible floodplain mapping. Chu and Steinman (2009) "
    "contributed the joint event\u2013continuous calibration strategy described in Section 2.2. In China, Oleyiblo and Li "
    "(2010) applied HEC-HMS with HEC-GeoHMS terrain preprocessing to the Misai and Wan\u2019an catchments and reported "
    "determination coefficients above 0.9 for all simulated flood events, with peak-discharge errors within acceptable "
    "forecasting limits, concluding that the model was suitable for flood forecasting in humid subtropical catchments."
)
body(
    "In South Asia, Meenu et al. (2013) used HEC-HMS 3.4 with the statistical downscaling model SDSM to assess the "
    "hydrologic impacts of HadCM3 A2 and B2 climate scenarios on the Tunga\u2013Bhadra basin in India, projecting increased "
    "precipitation and runoff for future periods; their study illustrates the use of HEC-HMS as the hydrological engine of "
    "climate-impact chains, an application later replicated in Bangladesh. Halwatura and Najim (2013) provided one of the "
    "most cited tropical calibration studies, comparing loss and transform method combinations for the Attanagalu Oya "
    "catchment in Sri Lanka and recommending the Snyder unit hydrograph with deficit-and-constant losses; their explicit "
    "message\u2014that method combinations must be tested rather than assumed\u2014has shaped subsequent practice. In Africa, "
    "Tassew et al. (2019) calibrated an event-based model of the 1,609 km\u00b2 Gilgel Abay catchment in the Lake Tana basin "
    "of Ethiopia using the SCS-CN, SCS unit hydrograph, and Muskingum methods, achieving a Nash\u2013Sutcliffe efficiency "
    "(NSE) of 0.884 and a coefficient of determination of 0.925 during validation, with sensitivity analysis identifying "
    "the curve number as the controlling parameter. Aliye et al. (2020) compared HEC-HMS and the Soil and Water Assessment "
    "Tool (SWAT) for a data-scarce catchment of the Ethiopian Rift Valley Lakes basin, exemplifying a growing comparative "
    "literature in which HEC-HMS holds its own against more heavily parameterized models. In Europe, Gilewski and Nawalany "
    "(2018) used HEC-HMS as the testbed for their precipitation-product intercomparison discussed above."
)
body(
    "Three consistent lessons emerge from this international experience. First, the combination of SCS-CN loss, SCS unit "
    "hydrograph transform, and Muskingum routing constitutes the de facto default configuration, and it generally performs "
    "satisfactorily in monsoonal and semi-arid regimes, but documented exceptions (Halwatura & Najim, 2013) forbid uncritical "
    "adoption. Second, the curve number is almost invariably the most sensitive parameter, so its spatial estimation and "
    "calibration bounds dominate model quality (Tassew et al., 2019). Third, HEC-HMS performance is limited less by its "
    "algorithms than by input data quality\u2014precipitation above all\u2014which motivates the careful forcing-data strategy "
    "reviewed in Section 2.4.3 (Gilewski & Nawalany, 2018)."
)

# ----------------------------------------------------------------- Table 2.1
tcap = doc.add_paragraph()
r = tcap.add_run("Table 2.1")
r.bold = True
tcap2 = doc.add_paragraph()
r = tcap2.add_run("Selected Applications of HEC-HMS Reviewed in This Chapter")
r.italic = True

rows = [
    ("Study", "Location / basin", "Principal methods", "Reported performance"),
    ("Knebl et al. (2005)", "San Antonio River basin, USA (~10,000 km\u00b2)",
     "NEXRAD rainfall; HEC-HMS + HEC-RAS coupling", "Calibrated discharge at 12 sub-basins; credible flood mapping"),
    ("Chu & Steinman (2009)", "Mona Lake watershed, Michigan, USA",
     "SCS-CN (event); SMA (continuous)", "Event-calibrated parameters improved continuous simulation"),
    ("Oleyiblo & Li (2010)", "Misai and Wan\u2019an catchments, China",
     "HEC-GeoHMS preprocessing; event simulation", "R\u00b2 > 0.9 for all events; acceptable peak errors"),
    ("Halwatura & Najim (2013)", "Attanagalu Oya catchment, Sri Lanka",
     "Snyder vs. Clark UH; SCS-CN vs. deficit-constant", "Snyder UH + deficit-constant most reliable; SCS-CN poor"),
    ("Meenu et al. (2013)", "Tunga\u2013Bhadra basin, India",
     "HEC-HMS 3.4 + SDSM downscaling (HadCM3 A2/B2)", "Projected increasing precipitation and runoff"),
    ("Gilewski & Nawalany (2018)", "Upper Skawa catchment, Poland",
     "Gauge vs. radar vs. IMERG forcing; SCS-CN", "Adjusted radar and IMERG most reliable forcings"),
    ("Tassew et al. (2019)", "Gilgel Abay, Lake Tana basin, Ethiopia (1,609 km\u00b2)",
     "SCS-CN; SCS UH; Muskingum", "NSE = 0.884; R\u00b2 = 0.925 (validation); CN most sensitive"),
    ("Aliye et al. (2020)", "Rift Valley Lakes basin, Ethiopia",
     "HEC-HMS vs. SWAT comparison", "Both models applicable in data-scarce region"),
    ("S. Haque et al. (2020)", "Brahmaputra River basin (at Bahadurabad), Bangladesh",
     "Continuous simulation; MUSLE sediment; RCP8.5", "NSE = 0.65 (cal.), 0.54 (val.); satisfactory"),
    ("Nujhat et al. (2024)", "Gumti River basin, Bangladesh",
     "SCS-CN; Muskingum; SRTM delineation", "R\u00b2 = 0.64 (cal.), 0.68 (val.); PBIAS very good"),
    ("M. B. Haque et al. (2024)", "Halda River catchment, Bangladesh",
     "SCS-CN optimized against SWAT-derived values", "NSE = 0.72 (cal.), 0.82 (val.); baseflow underestimated"),
    ("Moniruzzaman & Mahalder (2026)", "Atrai\u2013Karatoa River basin, Bangladesh",
     "HEC-HMS vs. SWAT comparison", "HEC-HMS R\u00b2 = 0.70 (cal.), 0.56 (val.); best for medium flows"),
]

table = doc.add_table(rows=len(rows), cols=4)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER
widths = [Inches(1.5), Inches(1.7), Inches(1.7), Inches(1.6)]
for i, row_data in enumerate(rows):
    for j, cell_text in enumerate(row_data):
        cell = table.cell(i, j)
        cell.width = widths[j]
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        run = p.add_run(cell_text)
        run.font.size = Pt(10)
        run.font.name = FONT
        if i == 0:
            run.bold = True

note = doc.add_paragraph()
r = note.add_run(
    "Note. NSE = Nash\u2013Sutcliffe efficiency; R\u00b2 = coefficient of determination; CN = curve number; "
    "UH = unit hydrograph; SMA = soil moisture accounting; cal. = calibration; val. = validation."
)
r.italic = True
r.font.size = Pt(10)

h2("2.6 Hydrological Modeling Experience in Bangladesh")
body(
    "Bangladesh occupies the lowermost reach of the Ganges\u2013Brahmaputra\u2013Meghna system, and its flood geography\u2014in "
    "which normal-season inundation of floodplains and beels is agriculturally essential while abnormal floods are "
    "catastrophic\u2014was authoritatively described by Brammer (1990) in his analysis of the 1987 and 1988 floods. Against "
    "this backdrop, a substantial national modeling capability has developed around institutional platforms, while the "
    "academic literature has increasingly adopted HEC-HMS for basin-scale rainfall\u2013runoff studies because of its zero "
    "cost and modest data demands."
)
body(
    "Several recent applications define the state of practice. S. Haque et al. (2020) developed a continuous HEC-HMS model "
    "of the poorly gauged Brahmaputra basin, calibrated at Bahadurabad against daily runoff for 1983\u20131996 (NSE = 0.65) "
    "and validated for 1997\u20132010 (NSE = 0.54), and extended it with the Modified Universal Soil Loss Equation and "
    "Engelund\u2013Hansen sediment routing to project sediment-load increases of 34%, 67%, and 115% by the 2020s, 2050s, and "
    "2080s under the RCP8.5 scenario. Nujhat et al. (2024) calibrated and validated an HEC-HMS model of the Gumti River "
    "basin using 2019\u20132021 records with SRTM-based delineation, Muskingum routing, and SCS-CN losses, obtaining "
    "coefficients of determination of 0.64 and 0.68 for calibration and validation respectively, with percent bias in the "
    "very good range, and recommended the calibrated model as a planning tool for flood prediction. M. B. Haque et al. "
    "(2024) modeled the ecologically important Halda catchment, optimizing curve numbers against values derived from a "
    "companion SWAT application and achieving NSE values of 0.72 and 0.82 in calibration and validation, while candidly "
    "documenting baseflow underestimation. Moniruzzaman and Mahalder (2026) compared HEC-HMS and SWAT for the "
    "Atrai\u2013Karatoa basin of northern Bangladesh\u2014a basin hydrologically contiguous with the greater Chalan Beel "
    "floodplain to which the Pabna beel systems belong\u2014finding that SWAT better captured high flows while HEC-HMS was "
    "more accurate for medium flows, with HEC-HMS achieving R\u00b2 of 0.70 and 0.56 in calibration and validation."
)
body(
    "Complementary studies using other models complete the picture. Raihan et al. (2020) simulated streamflow of the Upper "
    "Halda basin with SWAT (R\u00b2 = 0.80, NSE = 0.71), identifying groundwater delay, baseflow recession, and curve number "
    "as the most sensitive parameters, and emphasizing the difficulty of capturing rainfall variability from a single gauge. "
    "Akter and Ali (2012) assessed environmental flow requirements of the Halda River, illustrating the ecological "
    "motivation of much Bangladeshi hydrological modeling. Datta et al. (2022) contributed the DEM-sensitivity analysis of "
    "watershed delineation reviewed in Section 2.4.1."
)
body(
    "Taken together, the Bangladeshi HEC-HMS literature exhibits a clear spatial bias: published applications concentrate "
    "on large transboundary basins (Brahmaputra), flashy piedmont and hill catchments (Halda, Gumti), and the "
    "Atrai\u2013Karatoa corridor. The moribund distributary systems of the Ganges right bank\u2014of which the Ichamoti River "
    "of Pabna is a prominent example\u2014and their associated beel catchments have not, to the author\u2019s knowledge, been "
    "the subject of any published HEC-HMS rainfall\u2013runoff study. It is precisely in such systems, where local rainfall "
    "and drainage congestion rather than upstream river floods govern inundation, that runoff quantification is most needed "
    "for drainage design, yet where gauging is thinnest. This gap frames the contribution of the present thesis."
)

h2("2.7 Beel and Floodplain-Wetland Hydrology and the Ichamoti\u2013Gumai Beel Context")
body(
    "Wetland hydrology differs fundamentally from upland catchment hydrology: the hydroperiod\u2014the seasonal pattern of "
    "water-level rise and fall\u2014is the master variable controlling wetland ecology and land use, and it integrates "
    "precipitation, surface inflow and outflow, groundwater exchange, and evapotranspiration in a shallow storage whose "
    "area\u2013volume relationship is highly nonlinear (Mitsch & Gosselink, 2015). In the floodplains of Bangladesh, the "
    "characteristic wetland landform is the beel, a saucer-shaped depression on the floodplain that retains water "
    "perennially or seasonally and is filled both by direct rainfall-runoff from its local catchment and by spill from "
    "adjacent rivers during the monsoon (Brammer, 1990). Agriculture, capture fisheries, and settlement in beel areas are "
    "finely adapted to the normal hydroperiod, so that both deficient drainage (prolonged waterlogging) and excessive "
    "drainage (loss of dry-season water) constitute hazards (Brammer, 1990; Mitsch & Gosselink, 2015)."
)
body(
    "The hydrology of the beels of southwestern and northwestern Bangladesh cannot be understood apart from two "
    "generations of anthropogenic intervention. The first is the diversion of Ganges flow at the Farakka Barrage, "
    "commissioned in 1975, which significantly reduced dry-season discharge in the distributaries of the Bangladeshi "
    "Ganges. Mirza (1998) demonstrated statistically significant post-Farakka declines in dry-season flows of the Ganges "
    "and its Gorai distributary, with consequent accelerated sedimentation of distributary channels, and Gain and Giupponi "
    "(2014), applying the range-of-variability approach to twenty-two hydrologic indicators, showed that post-Farakka flows "
    "persistently failed pre-Farakka threshold conditions for both annual minima and maxima. Reduced parent-river flows "
    "starve distributary offtakes of the sediment-flushing discharges that keep them open, initiating the progressive "
    "siltation and eventual hydraulic disconnection of distributaries such as the Ichamoti. The second generation of "
    "intervention comprises the flood control, drainage, and irrigation (FCD/FCDI) projects constructed from the 1960s "
    "onward, including the Pabna Irrigation and Rural Development Project that encloses much of the study region. "
    "Thompson and Sultana (1996), evaluating the distributional impacts of such projects, documented that embankments "
    "frequently created internal drainage congestion\u2014to the point that people living inside protected areas deliberately "
    "cut embankments to release ponded water\u2014and that in four of five projects studied, flood losses inside protected "
    "areas during the 1988 flood exceeded those in adjacent unprotected areas. These findings establish drainage "
    "congestion, rather than riverine flooding alone, as the characteristic water hazard of embanked floodplain interiors."
)
body(
    "The Ichamoti River of Pabna exemplifies the resulting condition. Originating from the Padma near Shibrampur in Pabna "
    "Sadar and following a course of approximately 82\u201384 km through Pabna town toward the Hurasagar system in Bera "
    "upazila, the river has lost its connection with its parent rivers and is now widely described as dead or dying, its "
    "channel constricted by siltation, encroachment, solid-waste dumping, and aquatic weed infestation (\u201cIchhamati Now "
    "a Trickle,\u201d 2019; \u201cTk 1,554cr Project,\u201d 2024). Within Pabna municipality, the river bisects the town but "
    "no longer functions as an effective drainage outfall: Parvez et al. (2021) documented that the Ichamoti is "
    "moribund\u2014choked with water weeds and sediment\u2014and that parts of the municipality are subject to inundation "
    "during and after heavy rainfall, with residents identifying blocked and undersized drains, absent maintenance, and the "
    "unplanned drainage network as leading causes of waterlogging. Recognizing this condition, the Government of Bangladesh "
    "has undertaken a major rejuvenation program: a Tk 1,554.90 crore project implemented by the Bangladesh Water "
    "Development Board, involving dredging of a 33.77 km river stretch together with connecting canals including the "
    "Bharara channel, is intended to restore the Ichamoti\u2019s connectivity with the Padma and Jamuna systems "
    "(\u201cTk 1,554cr Project,\u201d 2024)."
)
body(
    "The Gumai Beel, the focus of the present study, is one of the beel systems of the Pabna floodplain drained through "
    "the Ichamoti corridor. Its inundation behavior is governed by the interaction of three controls that the reviewed "
    "literature identifies as characteristic of embanked distributary floodplains: local monsoon rainfall-runoff from the "
    "beel catchment; the conveyance capacity of the silted Ichamoti and its khals (canals), which sets the rate at which "
    "stored water can drain; and the stage of the receiving rivers, which can impose backwater limits on drainage. A "
    "rainfall\u2013runoff model of the beel catchment is the necessary first element of any quantitative analysis of this "
    "system: it supplies the inflow boundary condition for drainage design, for evaluation of the rejuvenation project\u2019s "
    "hydrological benefits, and for assessment of waterlogging risk under current and future rainfall regimes. The regional "
    "wetland literature reinforces the urgency of such quantification, as neighboring floodplain wetlands\u2014most "
    "prominently the greater Chalan Beel\u2014have experienced severe shrinkage and hydrological fragmentation under "
    "siltation, road and embankment construction, and land conversion (Brammer, 1990; Thompson & Sultana, 1996)."
)

h2("2.8 Model Calibration, Validation, and Performance Evaluation")
body(
    "Credible use of any rainfall\u2013runoff model requires a disciplined testing protocol. The canonical framework is the "
    "split-sample test of Klem\u0065\u0161 (1986): the observational record is divided so that the model is calibrated on one "
    "period and validated on an independent period, with more demanding differential tests (for example, calibration on wet "
    "years and validation on dry years) prescribed when the model is to be used under changed conditions. This protocol is "
    "followed, at least in its basic form, by essentially all of the applications reviewed above (e.g., Nujhat et al., "
    "2024; Oleyiblo & Li, 2010; S. Haque et al., 2020; Tassew et al., 2019)."
)
body(
    "Quantitative performance evaluation rests on a small set of statistics whose properties are well understood. The "
    "Nash\u2013Sutcliffe efficiency (Nash & Sutcliffe, 1970) measures the proportion of observed-flow variance explained by "
    "the model relative to the observed mean and remains the most widely reported metric, despite known sensitivities to "
    "peak flows and to the variance of the observation period. Gupta et al. (2009) decomposed the NSE into correlation, "
    "bias, and variability components and proposed the Kling\u2013Gupta efficiency as a more diagnostically transparent "
    "alternative, which is increasingly reported alongside NSE. For applied watershed modeling, the guidelines of Moriasi "
    "et al. (2007) have become the de facto standard: model performance for streamflow may be judged satisfactory when "
    "NSE exceeds 0.50, the RMSE\u2013observations standard deviation ratio (RSR) is at most 0.70, and percent bias (PBIAS) "
    "is within \u00b125%, with stricter thresholds for good and very good ratings. Moriasi et al. (2015) subsequently "
    "refined these ratings by constituent, time step, and model type, and recommended that graphical evaluation of "
    "hydrographs accompany all statistical assessment. The Bangladeshi applications reviewed in Section 2.6 adopt these "
    "thresholds explicitly, and the present study will do likewise, reporting NSE, R\u00b2, RSR, and PBIAS for both "
    "calibration and validation periods."
)
body(
    "Calibration itself may be manual, automated, or hybrid. HEC-HMS provides univariate-gradient and Nelder\u2013Mead "
    "search algorithms with a choice of objective functions (USACE, 2000), and reviewed applications span the full range "
    "of practice: manual calibration guided by physical reasoning (Knebl et al., 2005; Nujhat et al., 2024), automated "
    "optimization (Tassew et al., 2019), and staged strategies in which event calibration informs continuous simulation "
    "(Chu & Steinman, 2009). Sensitivity analysis should precede calibration to concentrate effort on influential "
    "parameters; the consistent finding that curve number, lag time, and Muskingum K dominate model response (Tassew et "
    "al., 2019) provides a defensible starting parameter set for the present study. Finally, the uncertainty literature "
    "cautions against over-interpretation of any single calibrated parameter set. Beven and Binley (1992) introduced the "
    "generalized likelihood uncertainty estimation framework on the premise that many parameter combinations may simulate "
    "observations equally well, and Beven (2006) elevated this observation into the equifinality thesis, arguing that "
    "environmental models should be evaluated as sets of acceptable simulators rather than as single optima. While a full "
    "uncertainty analysis lies beyond the scope of most application studies, acknowledging parameter equifinality\u2014for "
    "example, by reporting calibrated parameter ranges and testing alternative method combinations\u2014is now regarded as "
    "good practice, and the methodology of this thesis incorporates that principle."
)

h2("2.9 Synthesis and Research Gap")
body(
    "The reviewed literature supports four conclusions that together define the position of the present study. First, "
    "conceptual semi-distributed modeling with HEC-HMS is a mature, extensively validated approach whose default method "
    "combination (SCS-CN loss, SCS unit hydrograph transform, Muskingum routing) performs satisfactorily across monsoonal "
    "climates, provided that method choices are tested locally and that curve number estimation receives particular care "
    "(Halwatura & Najim, 2013; Tassew et al., 2019). Second, the data infrastructure required for such modeling\u2014global "
    "DEMs, satellite LULC, and blended precipitation products\u2014is available for Bangladesh, but flat deltaic terrain "
    "imposes recognized hazards on automated delineation that must be managed deliberately (Datta et al., 2022; Farr et "
    "al., 2007). Third, published Bangladeshi applications of HEC-HMS cluster in large transboundary and piedmont basins "
    "(M. B. Haque et al., 2024; Moniruzzaman & Mahalder, 2026; Nujhat et al., 2024; S. Haque et al., 2020); no published "
    "study has modeled rainfall\u2013runoff in the beel catchments of the moribund Ganges distributaries, of which the "
    "Ichamoti\u2013Gumai Beel system is representative. Fourth, the hydrological problem of such systems is distinctive: "
    "inundation is governed by local runoff accumulating behind congested drainage rather than by riverine flood waves "
    "(Parvez et al., 2021; Thompson & Sultana, 1996), while the long-term decline of the Ichamoti reflects basin-scale "
    "drivers, including post-Farakka flow reduction, that are well documented but have never been connected to a "
    "quantitative runoff model of the beel itself (Gain & Giupponi, 2014; Mirza, 1998)."
)
body(
    "The research gap addressed by this thesis follows directly. A calibrated and validated HEC-HMS rainfall\u2013runoff "
    "model of the Gumai Beel catchment will provide the first quantitative estimate of the runoff volumes and hydrograph "
    "dynamics that the Ichamoti drainage corridor must convey, thereby supplying the hydrological foundation for drainage "
    "design and for evaluation of the ongoing river rejuvenation program. Methodologically, the study extends the "
    "Bangladeshi HEC-HMS literature into an ultra-flat, wetland-dominated, data-scarce environment in which DEM "
    "limitations, seasonal land-cover alternation, and baseflow-storage behavior pose challenges identified but not "
    "resolved in previous work. The methods adopted in Chapter 3\u2014comparative testing of loss and transform methods, "
    "field-verified delineation, gauge-primary forcing with reanalysis gap filling, split-sample testing, and performance "
    "evaluation against the Moriasi et al. (2007, 2015) criteria\u2014are each grounded in the literature reviewed in this "
    "chapter."
)

# ================================================================ REFERENCES
page_break()
h1("References")

references = [
    "Akter, A., & Ali, M. H. (2012). Environmental flow requirements assessment in the Halda River, Bangladesh. "
    "Hydrological Sciences Journal, 57(2), 326\u2013343. https://doi.org/10.1080/02626667.2011.644242",

    "Aliye, M. A., Aga, A. O., Tadesse, T., & Yohannes, P. (2020). Evaluating the performance of HEC-HMS and SWAT "
    "hydrological models in simulating the rainfall-runoff process for data scarce region of Ethiopian Rift Valley Lake "
    "Basin. Open Journal of Modern Hydrology, 10(4), 105\u2013122. https://doi.org/10.4236/ojmh.2020.104007",

    "Beven, K. (2006). A manifesto for the equifinality thesis. Journal of Hydrology, 320(1\u20132), 18\u201336. "
    "https://doi.org/10.1016/j.jhydrol.2005.07.007",

    "Beven, K. J. (2012). Rainfall-runoff modelling: The primer (2nd ed.). Wiley-Blackwell. "
    "https://doi.org/10.1002/9781119951001",

    "Beven, K., & Binley, A. (1992). The future of distributed models: Model calibration and uncertainty prediction. "
    "Hydrological Processes, 6(3), 279\u2013298. https://doi.org/10.1002/hyp.3360060305",

    "Brammer, H. (1990). Floods in Bangladesh: Geographical background to the 1987 and 1988 floods. The Geographical "
    "Journal, 156(1), 12\u201322. https://doi.org/10.2307/635431",

    "Chow, V. T., Maidment, D. R., & Mays, L. W. (1988). Applied hydrology. McGraw-Hill.",

    "Chu, X., & Steinman, A. (2009). Event and continuous hydrologic modeling with HEC-HMS. Journal of Irrigation and "
    "Drainage Engineering, 135(1), 119\u2013124. https://doi.org/10.1061/(ASCE)0733-9437(2009)135:1(119)",

    "Datta, S., Karmakar, S., Mezbahuddin, S., Chaudhary, B. S., Hossain, M. M., Hoque, M. E., Abdullah-Al-Mamun, M. M., "
    "& Baul, T. K. (2022). The limits of watershed delineation: Implications of different DEMs, DEM resolutions, and area "
    "threshold values. Hydrology Research, 53(8), 1047\u20131062. https://doi.org/10.2166/nh.2022.126",

    "Devia, G. K., Ganasri, B. P., & Dwarakish, G. S. (2015). A review on hydrological models. Aquatic Procedia, 4, "
    "1001\u20131007. https://doi.org/10.1016/j.aqpro.2015.02.126",

    "Farr, T. G., Rosen, P. A., Caro, E., Crippen, R., Duren, R., Hensley, S., Kobrick, M., Paller, M., Rodriguez, E., "
    "Roth, L., Seal, D., Shaffer, S., Shimada, J., Umland, J., Werner, M., Oskin, M., Burbank, D., & Alsdorf, D. (2007). "
    "The Shuttle Radar Topography Mission. Reviews of Geophysics, 45(2), RG2004. https://doi.org/10.1029/2005RG000183",

    "Fleming, M. J., & Doan, J. H. (2013). HEC-GeoHMS geospatial hydrologic modeling extension: User\u2019s manual "
    "(Version 10.1). U.S. Army Corps of Engineers, Hydrologic Engineering Center.",

    "Funk, C., Peterson, P., Landsfeld, M., Pedreros, D., Verdin, J., Shukla, S., Husak, G., Rowland, J., Harrison, L., "
    "Hoell, A., & Michaelsen, J. (2015). The climate hazards infrared precipitation with stations\u2014A new environmental "
    "record for monitoring extremes. Scientific Data, 2, Article 150066. https://doi.org/10.1038/sdata.2015.66",

    "Gain, A. K., & Giupponi, C. (2014). Impact of the Farakka Dam on thresholds of the hydrologic flow regime in the "
    "Lower Ganges River Basin (Bangladesh). Water, 6(8), 2501\u20132518. https://doi.org/10.3390/w6082501",

    "Gilewski, P., & Nawalany, M. (2018). Inter-comparison of rain-gauge, radar, and satellite (IMERG GPM) precipitation "
    "estimates performance for rainfall-runoff modeling in a mountainous catchment in Poland. Water, 10(11), Article 1665. "
    "https://doi.org/10.3390/w10111665",

    "Gupta, H. V., Kling, H., Yilmaz, K. K., & Martinez, G. F. (2009). Decomposition of the mean squared error and NSE "
    "performance criteria: Implications for improving hydrological modelling. Journal of Hydrology, 377(1\u20132), 80\u201391. "
    "https://doi.org/10.1016/j.jhydrol.2009.08.003",

    "Halwatura, D., & Najim, M. M. M. (2013). Application of the HEC-HMS model for runoff simulation in a tropical "
    "catchment. Environmental Modelling & Software, 46, 155\u2013162. https://doi.org/10.1016/j.envsoft.2013.03.006",

    "Haque, M. B., Karmakar, S., & Hossain, M. M. (2024). Rainfall-runoff modeling using the HEC-HMS flow modeling "
    "framework for the Halda River catchment, Bangladesh [Preprint]. Research Square. "
    "https://doi.org/10.21203/rs.3.rs-3824469/v1",

    "Haque, S., Ali, M. M., Islam, A. K. M. S., & Khan, M. J. U. (2020). Changes in flow and sediment load of poorly "
    "gauged Brahmaputra river basin under an extreme climate scenario. Journal of Water and Climate Change, 12(3), "
    "937\u2013954. https://doi.org/10.2166/wcc.2020.219",

    "Hersbach, H., Bell, B., Berrisford, P., Hirahara, S., Hor\u00e1nyi, A., Mu\u00f1oz-Sabater, J., Nicolas, J., Peubey, C., "
    "Radu, R., Schepers, D., Simmons, A., Soci, C., Abdalla, S., Abellan, X., Balsamo, G., Bechtold, P., Biavati, G., "
    "Bidlot, J., Bonavita, M., \u2026 Th\u00e9paut, J.-N. (2020). The ERA5 global reanalysis. Quarterly Journal of the Royal "
    "Meteorological Society, 146(730), 1999\u20132049. https://doi.org/10.1002/qj.3803",

    "Ichhamati now a trickle. (2019). The Daily Star. "
    "https://www.thedailystar.net/backpage/news/ichhamati-now-trickle-1828042",

    "Klem\u0065\u0161, V. (1986). Operational testing of hydrological simulation models. Hydrological Sciences Journal, 31(1), "
    "13\u201324. https://doi.org/10.1080/02626668609491024",

    "Knebl, M. R., Yang, Z.-L., Hutchison, K., & Maidment, D. R. (2005). Regional scale flood modeling using NEXRAD "
    "rainfall, GIS, and HEC-HMS/RAS: A case study for the San Antonio River Basin summer 2002 storm event. Journal of "
    "Environmental Management, 75(4), 325\u2013336. https://doi.org/10.1016/j.jenvman.2004.11.024",

    "Meenu, R., Rehana, S., & Mujumdar, P. P. (2013). Assessment of hydrologic impacts of climate change in "
    "Tunga\u2013Bhadra river basin, India with HEC-HMS and SDSM. Hydrological Processes, 27(11), 1572\u20131589. "
    "https://doi.org/10.1002/hyp.9220",

    "Mirza, M. M. Q. (1998). Diversion of the Ganges water at Farakka and its effects on salinity in Bangladesh. "
    "Environmental Management, 22(5), 711\u2013722. https://doi.org/10.1007/s002679900141",

    "Mishra, S. K., & Singh, V. P. (2003). Soil Conservation Service Curve Number (SCS-CN) methodology. Springer. "
    "https://doi.org/10.1007/978-94-017-0147-1",

    "Mitsch, W. J., & Gosselink, J. G. (2015). Wetlands (5th ed.). John Wiley & Sons.",

    "Moniruzzaman, M., & Mahalder, B. (2026). Assessing SWAT and HEC-HMS model efficiency for watershed management in "
    "the Atrai-Karatoa River Basin, Bangladesh. Evolving Earth, 4, Article 100136. "
    "https://doi.org/10.1016/j.eve.2026.100136",

    "Moriasi, D. N., Arnold, J. G., Van Liew, M. W., Bingner, R. L., Harmel, R. D., & Veith, T. L. (2007). Model "
    "evaluation guidelines for systematic quantification of accuracy in watershed simulations. Transactions of the ASABE, "
    "50(3), 885\u2013900. https://doi.org/10.13031/2013.23153",

    "Moriasi, D. N., Gitau, M. W., Pai, N., & Daggupati, P. (2015). Hydrologic and water quality models: Performance "
    "measures and evaluation criteria. Transactions of the ASABE, 58(6), 1763\u20131785. "
    "https://doi.org/10.13031/trans.58.10715",

    "Nash, J. E., & Sutcliffe, J. V. (1970). River flow forecasting through conceptual models part I\u2014A discussion of "
    "principles. Journal of Hydrology, 10(3), 282\u2013290. https://doi.org/10.1016/0022-1694(70)90255-6",

    "Nujhat, M., Rayhan, M., & Amin, M. K. (2024). Hydrological modelling and its implication in sustainable water "
    "resource management in Gumti River Basin in Bangladesh. International Journal of Sustainability in Energy and "
    "Environment, 1(2), 40\u201348.",

    "Oleyiblo, J. O., & Li, Z. (2010). Application of HEC-HMS for flood forecasting in Misai and Wan\u2019an catchments in "
    "China. Water Science and Engineering, 3(1), 14\u201322. https://doi.org/10.3882/j.issn.1674-2370.2010.01.002",

    "Parvez, M., Sadat, N., Tasnim, F., & Nejhum, I. J. (2021). Identifying the causes of waterlogging on people\u2019s "
    "perception towards a resilient community: A case study on Pabna Municipality, Bangladesh. Ecofeminism and Climate "
    "Change, 2(3), 110\u2013126. https://doi.org/10.1108/EFCC-11-2020-0033",

    "Ponce, V. M., & Hawkins, R. H. (1996). Runoff curve number: Has it reached maturity? Journal of Hydrologic "
    "Engineering, 1(1), 11\u201319. https://doi.org/10.1061/(ASCE)1084-0699(1996)1:1(11)",

    "Raihan, F., Beaumont, L. J., Maina, J., Saiful Islam, A. K. M., & Harrison, S. P. (2020). Simulating streamflow in "
    "the Upper Halda Basin of southeastern Bangladesh using SWAT model. Hydrological Sciences Journal, 65(1), 138\u2013151. "
    "https://doi.org/10.1080/02626667.2019.1682149",

    "Singh, V. P., & Woolhiser, D. A. (2002). Mathematical modeling of watershed hydrology. Journal of Hydrologic "
    "Engineering, 7(4), 270\u2013292. https://doi.org/10.1061/(ASCE)1084-0699(2002)7:4(270)",

    "Tassew, B. G., Belete, M. A., & Miegel, K. (2019). Application of HEC-HMS model for flow simulation in the Lake "
    "Tana Basin: The case of Gilgel Abay catchment, Upper Blue Nile Basin, Ethiopia. Hydrology, 6(1), Article 21. "
    "https://doi.org/10.3390/hydrology6010021",

    "Thompson, P. M., & Sultana, P. (1996). Distributional and social impacts of flood control in Bangladesh. The "
    "Geographical Journal, 162(1), 1\u201313. https://doi.org/10.2307/3060212",

    "Tk 1,554cr project to revive dying Ichamati. (2024). The Daily Star. "
    "https://www.thedailystar.net/news/bangladesh/news/tk-1554cr-project-revive-dying-ichamati-3555931",

    "U.S. Army Corps of Engineers. (2000). Hydrologic Modeling System HEC-HMS: Technical reference manual. Hydrologic "
    "Engineering Center.",

    "USDA Soil Conservation Service. (1985). National engineering handbook, Section 4: Hydrology. U.S. Department of "
    "Agriculture.",
]

for r_text in references:
    ref(r_text)

OUT = "thesis/Literature_Review_Gumai_Beel_Ichamoti_HEC-HMS.docx"
doc.save(OUT)

words = sum(len(p.text.split()) for p in doc.paragraphs)
print(f"Saved {OUT}")
print(f"Approximate word count (paragraphs only): {words}")
