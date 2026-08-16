"""Generate the MS-thesis literature review chapter as an APA-7 formatted .docx.

Revision 2: all cited works are from 2019 onward. Every reference in the
bibliography was verified against publisher/indexing records (DOI, journal,
volume, issue, pages) before inclusion.
"""

from docx import Document
from docx.shared import Pt, Inches
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
    "This chapter reviews the recent body of scientific literature\u2014published from 2019 onward\u2014that underpins the "
    "present study, which develops a rainfall\u2013runoff model of the Gumai Beel catchment within the Ichamoti River system "
    "of Pabna District, Bangladesh, using the Hydrologic Engineering Center\u2019s Hydrologic Modeling System (HEC-HMS). "
    "Restricting the review to the most recent literature serves two purposes: it ensures that the methodological choices of "
    "this thesis reflect the current state of practice rather than superseded conventions, and it captures the newest "
    "generation of data products, evaluation standards, and regional studies that older reviews necessarily omit. The review "
    "is organized to move from the general to the particular. It begins with the conceptual foundations of rainfall\u2013runoff "
    "modeling and the contemporary debate on model selection, and then examines the structure of the HEC-HMS framework and the "
    "scientific basis of its principal computational methods. Subsequent sections review the role of geographic information "
    "systems (GIS) and remote sensing in supplying model inputs, survey recent applications of HEC-HMS across diverse "
    "hydro-climatic regions, and narrow the focus to hydrological modeling experience in Bangladesh. Because the study area is "
    "not an ordinary upland catchment but a low-lying floodplain wetland (beel) drained by a moribund distributary of the "
    "Padma River, a dedicated section reviews the hydrology of beel and floodplain-wetland systems, the anthropogenic decline "
    "of the Ichamoti River, and the chronic drainage congestion of the Pabna region. The chapter then reviews current "
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

h2("2.2 Rainfall\u2013Runoff Processes and the Selection of Hydrological Models")
body(
    "The transformation of rainfall into streamflow remains the central problem of applied hydrology. Precipitation falling "
    "on a catchment is partitioned among interception, depression storage, infiltration, evapotranspiration, soil-moisture "
    "replenishment, and surface runoff, and the runoff component is subsequently translated and attenuated as it travels over "
    "hillslopes and through channel networks to the catchment outlet. Because direct measurement of each of these processes at "
    "the catchment scale is impossible, hydrologists rely on mathematical models that represent the catchment as a system "
    "converting an input hyetograph into an output hydrograph. The contemporary modeling landscape is characterized by an "
    "overwhelming diversity of such models. Horton et al. (2022), in a systematic review of the drivers of model "
    "diversification, showed that the coexistence of empirical, conceptual, and physically based models\u2014and of lumped, "
    "semi-distributed, and fully distributed spatial structures\u2014reflects not only genuine differences in application "
    "context but also institutional habit, and they observed that the motivations for selecting a particular model are rarely "
    "stated explicitly in published studies. Their central recommendation, that model choice be justified by the adequacy of "
    "the model for the landscape, the data, and the purpose at hand, frames the model-selection reasoning of this thesis."
)
body(
    "For data-scarce developing-country catchments, the practical comparison is usually between conceptual semi-distributed "
    "systems such as HEC-HMS and more heavily parameterized ecohydrological models such as the Soil and Water Assessment "
    "Tool (SWAT). Recent head-to-head evaluations illuminate the trade-off. Aliye et al. (2020), comparing the two models for "
    "a data-scarce catchment of the Ethiopian Rift Valley Lakes basin, found both capable of credible streamflow simulation. "
    "Moniruzzaman and Mahalder (2026), performing the same comparison for the Atrai\u2013Karatoa basin of northern Bangladesh, "
    "found that SWAT better reproduced high flows while HEC-HMS was more accurate for medium flows, with both models "
    "performing satisfactorily overall. These results indicate that HEC-HMS remains a defensible choice where the modeling "
    "objective centers on runoff-hydrograph generation rather than on sediment, nutrient, or land-management simulation, and "
    "where its far smaller parameter set suits the available data."
)
body(
    "A further operational distinction separates event-based from continuous simulation. Event-based modeling simulates the "
    "catchment response to individual storms and is appropriate for design-flood estimation and flood forecasting; recent "
    "applications include the daily event-oriented modeling of the Al-Adhaim catchment in Iraq (Hamdan et al., 2021) and the "
    "flood-risk simulations of the Gumara River in Ethiopia (Admas et al., 2025). Continuous modeling, by contrast, tracks "
    "moisture accounting through wet and dry periods over months or years using algorithms such as the soil moisture "
    "accounting (SMA) loss method, as in the multi-year simulation of the Brahmaputra basin by Jawad (2024) and the "
    "climate-scenario modeling of the same basin by S. Haque et al. (2020). The inundation behavior of the Gumai Beel is "
    "governed both by individual monsoon storm events and by the strongly seasonal water balance, so the present study "
    "requires a modeling system capable of both modes, a requirement that HEC-HMS satisfies within a single framework "
    "(U.S. Army Corps of Engineers [USACE], n.d.)."
)

h2("2.3 The HEC-HMS Modeling Framework")

h3("2.3.1 Origin and Structure")
body(
    "HEC-HMS was developed by the Hydrologic Engineering Center of the U.S. Army Corps of Engineers to simulate the complete "
    "precipitation\u2013runoff process of dendritic watershed systems (USACE, n.d.). A HEC-HMS project comprises three "
    "principal components: a basin model, which describes the physical catchment as a network of sub-basins, reaches, "
    "junctions, reservoirs, diversions, sources, and sinks; a meteorological model, which assigns precipitation and "
    "evapotranspiration boundary conditions to the sub-basins; and control specifications, which define the simulation window "
    "and time step (USACE, n.d.). Within each sub-basin, the user selects one method from each of several interchangeable "
    "libraries\u2014canopy and surface storage, loss (infiltration), direct-runoff transform, and baseflow\u2014while each "
    "reach is assigned a channel-routing method. This modular architecture is the principal reason for the model\u2019s "
    "versatility: the same software can be configured as a simple event-based lumped model or as a continuous, "
    "semi-distributed model. The program is in the public domain, is supported by extensive online documentation, and "
    "integrates with GIS-based terrain preprocessing, all of which explain its wide adoption in regions where commercial "
    "modeling systems are unaffordable (Hamdan et al., 2021; Tassew et al., 2019)."
)

h3("2.3.2 Loss Methods and the SCS Curve Number")
body(
    "The loss model determines how much of the incident rainfall infiltrates or is otherwise abstracted, and therefore how "
    "much becomes precipitation excess available for direct runoff. The most widely used loss method in recent HEC-HMS "
    "applications remains the Soil Conservation Service Curve Number (SCS-CN) method, which condenses the combined influence "
    "of soil hydrologic group, land use, surface condition, and antecedent moisture into a single dimensionless parameter, "
    "the curve number (CN). Soulis (2021), reviewing the current state of the method more than six decades after its "
    "introduction, attributed its enduring dominance to its simplicity, its well-documented and easily obtained inputs, and "
    "its direct linkage to mappable catchment properties, while cataloguing the challenges that remain active research "
    "topics: the appropriate value of the initial abstraction ratio, the treatment of watershed slope, the representation of "
    "antecedent moisture, and the transfer of tabulated CN values to regions and land covers beyond those for which the "
    "method was derived. Notably, several studies reviewed by Soulis (2021) found that the conventional initial abstraction "
    "ratio of 0.2 substantially overestimates initial losses, with locally calibrated values often below 0.05, a finding "
    "directly relevant to CN calibration bounds in the present study."
)
body(
    "HEC-HMS offers several alternatives to the SCS-CN approach, including the initial-and-constant and deficit-and-constant "
    "loss methods, the Green\u2013Ampt infiltration model, and, for continuous simulation, the five-layer SMA algorithm "
    "(USACE, n.d.). Recent empirical evidence indicates that the SCS-CN method performs well in monsoonal and semi-arid "
    "event simulation: Tassew et al. (2019) obtained very good performance with it in the Gilgel Abay catchment of Ethiopia, "
    "and Hamdan et al. (2021) reached the same conclusion for the Al-Adhaim catchment in Iraq. For continuous multi-year "
    "simulation, the SMA method is preferred, as in the Brahmaputra applications of Jawad (2024) and S. Haque et al. (2020). "
    "Equally consistent is the finding that the curve number is the single most sensitive parameter of the model: Tassew et "
    "al. (2019) identified CN as the controlling parameter in formal sensitivity analysis, and M. B. Haque et al. (2024) "
    "found the runoff\u2013rainfall relationship of the Halda catchment to be highly sensitive to sub-basin CN values, which "
    "depend directly on the land-use classification. These findings justify the comparative testing of loss methods and the "
    "careful CN estimation undertaken in the methodology of this thesis."
)

h3("2.3.3 Direct-Runoff Transform Methods")
body(
    "The transform model converts precipitation excess into a direct-runoff hydrograph at the sub-basin outlet. Most "
    "applications employ unit-hydrograph theory, for which HEC-HMS provides the SCS dimensionless unit hydrograph, "
    "parameterized by basin lag time; the Clark unit hydrograph, which combines a time\u2013area histogram with a "
    "linear-reservoir storage coefficient; and the Snyder unit hydrograph, among others (USACE, n.d.). The SCS unit "
    "hydrograph paired with SCS-CN losses constitutes the dominant configuration in the recent application literature across "
    "Africa, the Middle East, and South Asia (Hamdan et al., 2021; Nujhat et al., 2024; Tassew et al., 2019), while the "
    "Clark formulation has been preferred in some continuous large-basin applications (Jawad, 2024). For very flat "
    "catchments such as the study area, lag-time estimation deserves particular care, because empirical lag equations were "
    "developed predominantly for sloping terrain and misestimated lag propagates directly into errors in simulated peak "
    "timing; the calibration of time-of-concentration and storage parameters was accordingly a central element of the "
    "optimization strategies reported by Jawad (2024) and Admas et al. (2025)."
)

h3("2.3.4 Baseflow Methods")
body(
    "Baseflow representation is frequently the weakest element of event-oriented models, yet in floodplain-wetland systems "
    "the slow drainage component can dominate the recession limb and the dry-season water balance. HEC-HMS provides "
    "recession, bounded-recession, linear-reservoir, and constant-monthly baseflow methods (USACE, n.d.); the "
    "linear-reservoir method with two groundwater layers was adopted in the continuous Brahmaputra modeling of Jawad (2024). "
    "Recent Bangladeshi experience suggests that baseflow is a genuine difficulty: M. B. Haque et al. (2024), modeling the "
    "Halda River catchment, obtained satisfactory overall statistics but reported a poor match for the baseflow portion of "
    "the hydrograph during calibration, which they attributed to unrepresented groundwater\u2013surface water exchange, and "
    "they recommended coupling with a groundwater model as a route to improvement. Such findings caution against "
    "interpreting event-calibrated models as complete descriptions of low-flow behavior, a caution of particular force in "
    "beel environments where monsoon storage is released gradually through the post-monsoon season."
)

h3("2.3.5 Channel Routing and Hydraulic Coupling")
body(
    "Flow routing through reaches is available in HEC-HMS through the Muskingum, Muskingum\u2013Cunge, kinematic-wave, "
    "modified-Puls, and lag methods (USACE, n.d.). The Muskingum method, which represents a reach as a linear storage with "
    "travel-time parameter K and weighting parameter X, remains the most commonly adopted in applications comparable to the "
    "present study (Hamdan et al., 2021; Nujhat et al., 2024; Tassew et al., 2019). Its limitation\u2014shared by all "
    "hydrologic routing schemes\u2014is that it cannot represent backwater effects, flow reversal, or looped stage\u2013"
    "discharge relations, which are common in extremely flat deltaic channels. Where such hydraulic effects matter, the "
    "recent literature couples HEC-HMS with the HEC-RAS hydraulic model: Admas et al. (2025) drove HEC-RAS flood-risk "
    "simulations of the Gumara River floodplain in Ethiopia with HEC-HMS hydrographs to quantify inundation for return "
    "periods up to 100 years, and Jawad (2024) coupled the two models to map flood inundation along the lowermost 500 km of "
    "the Brahmaputra. An analogous coupling represents a natural extension of the present work, given that drainage of the "
    "Gumai Beel is partly controlled by water levels in the receiving Ichamoti channel rather than by catchment runoff alone."
)

h2("2.4 GIS, Remote Sensing, and Input Data for Hydrological Modeling")

h3("2.4.1 Digital Elevation Models and Watershed Delineation")
body(
    "Semi-distributed modeling begins with terrain analysis: delineation of sub-basins and stream networks from a digital "
    "elevation model (DEM), and extraction of physiographic parameters such as area, slope, and longest flow path. A "
    "growing family of freely available global DEMs\u2014SRTM, NASADEM, ASTER, AW3D30, MERIT, and TanDEM-X\u2014now competes "
    "for this role, and their quality differences are material. Uuemaa et al. (2020), in a systematic multi-region accuracy "
    "assessment against LiDAR references, found vertical accuracy to vary strongly among products and identified terrain "
    "slope as the dominant control on DEM error, with AW3D30 the most robust performer overall and NASADEM only marginally "
    "improving on SRTM. The reliability of automated delineation degrades precisely in the terrain type of the present "
    "study: low-relief floodplains. Datta et al. (2022), working on the Halda watershed in Bangladesh, showed systematically "
    "that delineation outcomes depend materially on the choice of DEM product, its spatial resolution, and the "
    "stream-definition area threshold, and that these choices propagate into sub-basin geometry and derived parameters. In "
    "flat deltaic terrain, where total relief may be only a few meters and anthropogenic features such as roads and "
    "embankments control actual flow paths, DEM vertical error can exceed the topographic signal, and the literature "
    "therefore recommends verification of automatically delineated drainage against field knowledge and hydrographic maps "
    "(Datta et al., 2022). This consideration is central to the delineation strategy adopted in Chapter 3 of this thesis."
)

h3("2.4.2 Land Use, Soils, and Curve Number Generation")
body(
    "Loss-model parameterization in ungauged or sparsely gauged basins rests on thematic mapping. The standard workflow "
    "intersects a land-use/land-cover (LULC) classification, commonly derived from Landsat or Sentinel-2 imagery, with a "
    "hydrologic soil group map derived from soil surveys or global databases, and assigns composite curve numbers to each "
    "sub-basin through standard lookup tables (Hamdan et al., 2021; Soulis, 2021). Applications across climates confirm "
    "both the practicality and the sensitivity of this workflow: Tassew et al. (2019) identified the curve number as the "
    "single most sensitive parameter of their model, and M. B. Haque et al. (2024) likewise found simulated runoff to be "
    "highly sensitive to sub-basin CN values, which depend directly on the LULC classification. Remote sensing also "
    "documents the pace of land-cover change in Bangladeshi wetland environments: Bhattacharjee et al. (2021), analyzing "
    "three decades of Landsat imagery over a northeastern wetland (haor) system, quantified substantial conversion of "
    "wetland and vegetation classes to agriculture and settlement. These findings imply that LULC currency and accuracy "
    "assessment deserve explicit attention in the present study, particularly because beel environments exhibit strong "
    "seasonal alternation between open water, cropped land, and marsh vegetation that a single-date classification cannot "
    "capture."
)

h3("2.4.3 Precipitation Inputs: Gauges, Reanalysis, and Satellite Products")
body(
    "Precipitation is the dominant control on simulated runoff, and its estimation is a major source of uncertainty where "
    "gauge networks are sparse. Two families of alternatives to gauge interpolation have matured over the past decade: "
    "satellite multi-sensor products, foremost the Integrated Multi-satellitE Retrievals for GPM (IMERG), and atmospheric "
    "reanalyses, foremost ERA5 (Hersbach et al., 2020). Pradhan et al. (2022), in a global systematic review of IMERG "
    "validation studies, concluded that the product reliably captures regional precipitation patterns and improves with "
    "every version, while performing better at monthly than at daily and sub-daily scales and retaining known weaknesses "
    "for extreme intensities and complex terrain\u2014weaknesses that matter for flood-oriented applications. In hydrological "
    "practice, satellite precipitation has proven usable even for operational simulation of very large, sparsely gauged "
    "basins: Jawad (2024) forced coupled HEC-HMS\u2013HEC-RAS models of the Brahmaputra with near-real-time IMERG and GSMaP "
    "products, gauge-corrected where possible, and obtained credible discharge and inundation simulations at Bahadurabad. "
    "For the present study, which can draw on Bangladesh Water Development Board and Bangladesh Meteorological Department "
    "gauges in and around Pabna, this literature supports a strategy of gauge-primary forcing with reanalysis or satellite "
    "products used for gap filling and consistency checking (Hersbach et al., 2020; Pradhan et al., 2022)."
)

h2("2.5 Recent Applications of HEC-HMS Across Diverse Hydro-Climatic Regions")
body(
    "The post-2019 application literature on HEC-HMS is extensive, and a selective review of methodologically instructive "
    "studies is presented here; Table 2.1 summarizes their essential characteristics. In Africa, Tassew et al. (2019) "
    "calibrated an event-based model of the 1,609 km\u00b2 Gilgel Abay catchment in the Lake Tana basin of Ethiopia using "
    "the SCS-CN, SCS unit hydrograph, and Muskingum methods, achieving a Nash\u2013Sutcliffe efficiency (NSE) of 0.884 and a "
    "coefficient of determination of 0.925, with sensitivity analysis identifying the curve number as the controlling "
    "parameter. Aliye et al. (2020) compared HEC-HMS and SWAT for a data-scarce catchment of the Ethiopian Rift Valley "
    "Lakes basin, finding both models applicable and thereby establishing HEC-HMS as a viable lightweight alternative to "
    "more heavily parameterized systems. Admas et al. (2025) extended the Ethiopian experience to integrated flood-risk "
    "analysis, coupling HEC-HMS design hydrographs with HEC-RAS hydraulics for the Gumara River floodplain and quantifying "
    "the inundation reduction achieved by dyke construction for return periods from 2 to 100 years."
)
body(
    "In the Middle East, Hamdan et al. (2021) developed a daily model of the semi-arid Al-Adhaim catchment in Iraq with "
    "HEC-GeoHMS preprocessing and the SCS-CN, SCS unit hydrograph, and Muskingum configuration, obtaining coefficients of "
    "determination near 0.9 in both calibration and verification and concluding that the model is suitable for reservoir "
    "inflow estimation. In South Asia, Jawad (2024) tested the frontier of data-scarce application by forcing a continuous "
    "SMA-based HEC-HMS model of the transboundary Brahmaputra basin entirely with near-real-time satellite precipitation, "
    "calibrating by automated univariate-gradient search and coupling the results to HEC-RAS for flood inundation mapping. "
    "Collectively, three consistent lessons emerge from this international experience. First, the combination of SCS-CN "
    "loss, SCS unit hydrograph transform, and Muskingum routing constitutes the de facto default configuration for "
    "event-scale application, and it performs satisfactorily across monsoonal and semi-arid regimes (Hamdan et al., 2021; "
    "Tassew et al., 2019). Second, the curve number is almost invariably the most sensitive parameter, so its spatial "
    "estimation and calibration bounds dominate model quality (M. B. Haque et al., 2024; Tassew et al., 2019). Third, "
    "HEC-HMS performance is limited less by its algorithms than by input data quality\u2014precipitation above all\u2014which "
    "motivates the careful forcing-data strategy reviewed in Section 2.4.3 (Jawad, 2024; Pradhan et al., 2022)."
)

# ----------------------------------------------------------------- Table 2.1
tcap = doc.add_paragraph()
r = tcap.add_run("Table 2.1")
r.bold = True
tcap2 = doc.add_paragraph()
r = tcap2.add_run("Selected Post-2019 Applications of HEC-HMS and Comparable Models Reviewed in This Chapter")
r.italic = True

rows = [
    ("Study", "Location / basin", "Principal methods", "Reported performance"),
    ("Tassew et al. (2019)", "Gilgel Abay, Lake Tana basin, Ethiopia (1,609 km\u00b2)",
     "SCS-CN; SCS UH; Muskingum", "NSE = 0.884; R\u00b2 = 0.925; CN most sensitive"),
    ("Aliye et al. (2020)", "Rift Valley Lakes basin, Ethiopia",
     "HEC-HMS vs. SWAT comparison", "Both models applicable in data-scarce region"),
    ("S. Haque et al. (2020)", "Brahmaputra basin (at Bahadurabad), Bangladesh",
     "Continuous simulation; MUSLE sediment; RCP8.5", "NSE = 0.65 (cal.), 0.54 (val.); satisfactory"),
    ("Raihan et al. (2020)", "Upper Halda basin, Bangladesh",
     "SWAT (comparison baseline)", "R\u00b2 = 0.80; NSE = 0.71; CN among most sensitive"),
    ("Hamdan et al. (2021)", "Al-Adhaim catchment, Iraq",
     "SCS-CN; SCS UH; Muskingum; HEC-GeoHMS", "R\u00b2 \u2248 0.90 (calibration and verification)"),
    ("Jawad (2024)", "Brahmaputra basin (transboundary)",
     "SMA continuous; Clark UH; linear-reservoir baseflow; IMERG/GSMaP forcing; HEC-RAS coupling",
     "Credible discharge and inundation simulation from satellite forcing"),
    ("Nujhat et al. (2024)", "Gumti River basin, Bangladesh",
     "SCS-CN; Muskingum; SRTM delineation", "R\u00b2 = 0.64 (cal.), 0.68 (val.); PBIAS very good"),
    ("M. B. Haque et al. (2024)", "Halda River catchment, Bangladesh",
     "SCS-CN optimized against SWAT-derived values", "NSE = 0.72 (cal.), 0.82 (val.); baseflow underestimated"),
    ("Admas et al. (2025)", "Gumara River, Upper Blue Nile basin, Ethiopia",
     "HEC-HMS + HEC-RAS flood-risk coupling", "Inundation quantified for 2\u2013100-year return periods"),
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
    "Bangladesh occupies the lowermost reach of the Ganges\u2013Brahmaputra\u2013Meghna system, in which normal-season "
    "inundation of floodplains and beels is agriculturally essential while abnormal floods and chronic waterlogging are "
    "hazards. Against this backdrop, the national academic literature has increasingly adopted HEC-HMS for basin-scale "
    "rainfall\u2013runoff studies because of its zero cost and modest data demands, and the post-2019 record now spans the "
    "country\u2019s principal basin types."
)
body(
    "Several recent applications define the state of practice. S. Haque et al. (2020) developed a continuous HEC-HMS model "
    "of the poorly gauged Brahmaputra basin, calibrated at Bahadurabad against daily runoff for 1983\u20131996 (NSE = 0.65) "
    "and validated for 1997\u20132010 (NSE = 0.54), and extended it with the Modified Universal Soil Loss Equation and "
    "Engelund\u2013Hansen sediment routing to project sediment-load increases of 34%, 67%, and 115% by the 2020s, 2050s, and "
    "2080s under the RCP8.5 scenario\u2014demonstrating the use of HEC-HMS as the hydrological engine of climate-impact "
    "chains. Nujhat et al. (2024) calibrated and validated a model of the Gumti River basin using 2019\u20132021 records with "
    "SRTM-based delineation, Muskingum routing, and SCS-CN losses, obtaining coefficients of determination of 0.64 and 0.68 "
    "for calibration and validation respectively, with percent bias in the very good range, and recommended the calibrated "
    "model as a planning tool for flood prediction. M. B. Haque et al. (2024) modeled the ecologically important Halda "
    "catchment, optimizing curve numbers against values derived from a companion SWAT application and achieving NSE values "
    "of 0.72 and 0.82 in calibration and validation, while candidly documenting baseflow underestimation. Moniruzzaman and "
    "Mahalder (2026) compared HEC-HMS and SWAT for the Atrai\u2013Karatoa basin of northern Bangladesh\u2014a basin "
    "hydrologically contiguous with the greater Chalan Beel floodplain to which the Pabna beel systems belong\u2014finding "
    "that SWAT better captured high flows while HEC-HMS was more accurate for medium flows, with HEC-HMS achieving "
    "R\u00b2 of 0.70 and 0.56 in calibration and validation."
)
body(
    "Complementary studies using other models complete the picture. Raihan et al. (2020) simulated streamflow of the Upper "
    "Halda basin with SWAT (R\u00b2 = 0.80, NSE = 0.71), identifying groundwater delay, baseflow recession, and curve number "
    "as the most sensitive parameters, and emphasizing the difficulty of capturing rainfall variability from a single gauge. "
    "Datta et al. (2022) contributed the DEM-sensitivity analysis of watershed delineation reviewed in Section 2.4.1, and "
    "Jawad (2024) demonstrated satellite-forced coupled hydrologic\u2013hydraulic modeling for the Bangladeshi reach of the "
    "Brahmaputra."
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
    "Wetland hydrology differs fundamentally from upland catchment hydrology: the water regime\u2014the seasonal pattern of "
    "water-level rise and fall\u2014is the master variable controlling wetland biogeochemistry, ecology, and land use, and "
    "altered hydrology is correspondingly identified as the principal pathway by which environmental change degrades wetland "
    "function (Salimi et al., 2021). In the floodplains of Bangladesh, the characteristic wetland landform is the beel, a "
    "saucer-shaped depression on the floodplain that retains water perennially or seasonally and is filled both by direct "
    "rainfall-runoff from its local catchment and by spill from adjacent rivers during the monsoon (Adnan et al., 2020). "
    "Agriculture, capture fisheries, and settlement in beel areas are finely adapted to the normal seasonal water regime, so "
    "that both deficient drainage (prolonged waterlogging) and excessive drainage (loss of dry-season water) constitute "
    "hazards, and remote-sensing studies of Bangladeshi wetland systems document rapid ongoing conversion of wetland classes "
    "to agriculture and settlement under such pressures (Bhattacharjee et al., 2021)."
)
body(
    "The hydrology of the beels of western Bangladesh cannot be understood apart from two generations of anthropogenic "
    "intervention. The first is the long-term decline of dry-season flows in the Ganges distributary network. Ali and Hasan "
    "(2022), analyzing Bangladesh Water Development Board discharge records for the Gorai\u2014the principal Ganges "
    "distributary and the regional analogue of the Ichamoti\u2014found that mean annual flow in 2000\u20132016 was about 13% "
    "lower than in 1984\u20131999, quantified a deficient-flow condition persisting from December to May, and estimated that "
    "the river now frequently fails to meet its environmental flow requirement of roughly 29% of mean annual flow. Reduced "
    "parent-river flows starve distributary offtakes of the sediment-flushing discharges that keep them open, initiating "
    "the progressive siltation and eventual hydraulic disconnection of distributaries such as the Ichamoti. The second "
    "generation of intervention comprises the embankment, polder, and flood control\u2013drainage\u2013irrigation "
    "infrastructure constructed from the 1960s onward, including the Pabna Irrigation and Rural Development Project that "
    "encloses much of the study region. The systemic consequence of such enclosure, documented most thoroughly for the "
    "embanked southwest delta, is internal drainage congestion: embankments disconnect the floodplain from the river "
    "network, sediment accretes in the riverbeds rather than on the floodplain, gravity drainage progressively fails, and "
    "monsoon runoff accumulates in the enclosed beels as pluvial flooding (Adnan et al., 2020). Adnan et al. (2020) showed "
    "that in the embanked southwest, a large majority of agricultural land now lies in flood-susceptible zones, and they "
    "evaluated controlled sediment reintroduction (tidal river management) in low-lying beels as a rehabilitation measure\u2014"
    "evidence that beel drainage problems are structural and regional rather than local anomalies."
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
    "stored water can drain; and the stage of the receiving rivers, which can impose backwater limits on drainage (Adnan "
    "et al., 2020; Parvez et al., 2021). A rainfall\u2013runoff model of the beel catchment is the necessary first element "
    "of any quantitative analysis of this system: it supplies the inflow boundary condition for drainage design, for "
    "evaluation of the rejuvenation project\u2019s hydrological benefits, and for assessment of waterlogging risk under "
    "current and future rainfall regimes. The regional wetland literature reinforces the urgency of such quantification, "
    "as neighboring floodplain wetlands\u2014most prominently the greater Chalan Beel system to which the Atrai\u2013Karatoa "
    "drainage belongs\u2014have experienced severe shrinkage and hydrological fragmentation under siltation, embankment "
    "construction, and land conversion (Bhattacharjee et al., 2021; Moniruzzaman & Mahalder, 2026)."
)

h2("2.8 Model Calibration, Validation, and Performance Evaluation")
body(
    "Credible use of any rainfall\u2013runoff model requires a disciplined testing protocol. The conventional framework is "
    "the split-sample test, in which the observational record is divided so that the model is calibrated on one period and "
    "validated on an independent period; essentially all of the applications reviewed above follow this protocol (Hamdan "
    "et al., 2021; Nujhat et al., 2024; S. Haque et al., 2020; Tassew et al., 2019). The framework itself, however, is "
    "under active re-examination. Shen et al. (2022), in a large-sample experiment spanning 463 catchments, two conceptual "
    "models, and 50 data-splitting schemes, found that the common practice of calibrating to older data and validating on "
    "newer data systematically degrades subsequent predictive performance, and that calibrating to all available data is "
    "the most robust strategy when the model will be used operationally. For a thesis context, where independent "
    "demonstration of predictive skill remains an examination requirement, these findings argue for reporting both a "
    "conventional split-sample evaluation and a final parameter set re-estimated on the full record\u2014the approach "
    "adopted in Chapter 3."
)
body(
    "Quantitative performance evaluation rests on a small set of statistics whose properties have been sharply clarified "
    "in recent years. The Nash\u2013Sutcliffe efficiency remains the most widely reported metric, with values above zero "
    "indicating improvement on the observed-mean benchmark and values above approximately 0.5 conventionally regarded as "
    "satisfactory for streamflow, as applied in the Bangladeshi and Ethiopian studies reviewed above (Moniruzzaman & "
    "Mahalder, 2026; Nujhat et al., 2024; Tassew et al., 2019). The Kling\u2013Gupta efficiency (KGE), which decomposes "
    "performance into correlation, bias, and variability components, is increasingly reported alongside NSE; Knoben et al. "
    "(2019) demonstrated, however, that KGE and NSE values are not directly comparable\u2014the observed-mean benchmark "
    "corresponds to KGE \u2248 \u22120.41 rather than zero\u2014and cautioned modelers against transferring NSE-based "
    "intuitions to KGE scores. Althoff and Rodrigues (2021) provided a systematic analysis of goodness-of-fit criteria for "
    "model calibration and evaluation, showing that the choice of objective function materially shapes the resulting "
    "parameter set and recommending that criteria be matched deliberately to the modeling purpose\u2014for example, "
    "peak-oriented criteria for flood design versus volume-oriented criteria for water-balance assessment. Percent bias "
    "and the RMSE\u2013observations standard deviation ratio complete the standard reporting set used in the recent "
    "Bangladeshi applications (M. B. Haque et al., 2024; Nujhat et al., 2024; S. Haque et al., 2020), and the present "
    "study will report all four statistics for both calibration and validation periods, accompanied by graphical "
    "hydrograph comparison."
)
body(
    "Calibration itself may be manual, automated, or hybrid. HEC-HMS provides univariate-gradient and simplex search "
    "algorithms with a choice of objective functions (USACE, n.d.), and reviewed applications span the full range of "
    "practice: manual calibration guided by physical reasoning (Nujhat et al., 2024), automated optimization (Jawad, 2024; "
    "Tassew et al., 2019), and staged strategies in which sensitivity analysis precedes calibration to concentrate effort "
    "on influential parameters (Admas et al., 2025). The consistent finding that curve number, lag time, and Muskingum K "
    "dominate model response (Tassew et al., 2019) provides a defensible starting parameter set for the present study. "
    "Finally, the recent metric literature carries an implicit warning about equifinality: because different objective "
    "functions select different acceptable parameter sets (Althoff & Rodrigues, 2021), and because aggregate scores can "
    "mask compensating errors in correlation, bias, and variability (Knoben et al., 2019), good practice now favors "
    "reporting calibrated parameter ranges, decomposed performance components, and alternative method combinations rather "
    "than a single nominally optimal model\u2014a principle incorporated in the methodology of this thesis."
)

h2("2.9 Synthesis and Research Gap")
body(
    "The reviewed literature supports four conclusions that together define the position of the present study. First, "
    "conceptual semi-distributed modeling with HEC-HMS is a mature, extensively validated approach whose default method "
    "combination (SCS-CN loss, SCS unit hydrograph transform, Muskingum routing) performs satisfactorily across monsoonal "
    "and semi-arid climates, provided that method choices are justified for the landscape and purpose at hand and that "
    "curve number estimation receives particular care (Hamdan et al., 2021; Horton et al., 2022; Soulis, 2021; Tassew et "
    "al., 2019). Second, the data infrastructure required for such modeling\u2014global DEMs, satellite land cover, and "
    "satellite or reanalysis precipitation\u2014is available for Bangladesh, but flat deltaic terrain imposes recognized "
    "hazards on automated delineation, and precipitation-product limitations for extremes must be managed deliberately "
    "(Datta et al., 2022; Pradhan et al., 2022; Uuemaa et al., 2020). Third, published Bangladeshi applications of HEC-HMS "
    "cluster in large transboundary and piedmont basins (Jawad, 2024; M. B. Haque et al., 2024; Moniruzzaman & Mahalder, "
    "2026; Nujhat et al., 2024; S. Haque et al., 2020); no published study has modeled rainfall\u2013runoff in the beel "
    "catchments of the moribund Ganges distributaries, of which the Ichamoti\u2013Gumai Beel system is representative. "
    "Fourth, the hydrological problem of such systems is distinctive: inundation is governed by local runoff accumulating "
    "behind congested drainage rather than by riverine flood waves (Adnan et al., 2020; Parvez et al., 2021), while the "
    "long-term decline of the distributary network reflects basin-scale flow reduction that is well quantified for the "
    "neighboring Gorai but has never been connected to a quantitative runoff model of the beel itself (Ali & Hasan, 2022)."
)
body(
    "The research gap addressed by this thesis follows directly. A calibrated and validated HEC-HMS rainfall\u2013runoff "
    "model of the Gumai Beel catchment will provide the first quantitative estimate of the runoff volumes and hydrograph "
    "dynamics that the Ichamoti drainage corridor must convey, thereby supplying the hydrological foundation for drainage "
    "design and for evaluation of the ongoing river rejuvenation program (\u201cTk 1,554cr Project,\u201d 2024). "
    "Methodologically, the study extends the Bangladeshi HEC-HMS literature into an ultra-flat, wetland-dominated, "
    "data-scarce environment in which DEM limitations, seasonal land-cover alternation, and baseflow-storage behavior pose "
    "challenges identified but not resolved in previous work (Datta et al., 2022; M. B. Haque et al., 2024). The methods "
    "adopted in Chapter 3\u2014comparative testing of loss and transform methods, field-verified delineation, gauge-primary "
    "forcing with reanalysis gap filling, split-sample testing informed by current calibration research, and performance "
    "evaluation using NSE, KGE, percent bias, and RSR with purpose-matched objective functions\u2014are each grounded in "
    "the literature reviewed in this chapter (Althoff & Rodrigues, 2021; Knoben et al., 2019; Shen et al., 2022)."
)

# ================================================================ REFERENCES
page_break()
h1("References")

references = [
    "Adnan, M. S. G., Talchabhadel, R., Nakagawa, H., & Hall, J. W. (2020). The potential of Tidal River Management for "
    "flood alleviation in South Western Bangladesh. Science of the Total Environment, 731, Article 138747. "
    "https://doi.org/10.1016/j.scitotenv.2020.138747",

    "Admas, M., Asrade, T. M., & Cherie, W. D. (2025). Application of the HEC-RAS and HEC-HMS models for flood risk "
    "analysis in the Gumara River, Upper Blue Nile Basin, Ethiopia. Advances in Meteorology, 2025, Article 5092932. "
    "https://doi.org/10.1155/adme/5092932",

    "Ali, M. S., & Hasan, M. M. (2022). Environmental flow assessment of Gorai River in Bangladesh: A comparative "
    "analysis of different hydrological methods. Heliyon, 8(7), Article e09857. "
    "https://doi.org/10.1016/j.heliyon.2022.e09857",

    "Aliye, M. A., Aga, A. O., Tadesse, T., & Yohannes, P. (2020). Evaluating the performance of HEC-HMS and SWAT "
    "hydrological models in simulating the rainfall-runoff process for data scarce region of Ethiopian Rift Valley Lake "
    "Basin. Open Journal of Modern Hydrology, 10(4), 105\u2013122. https://doi.org/10.4236/ojmh.2020.104007",

    "Althoff, D., & Rodrigues, L. N. (2021). Goodness-of-fit criteria for hydrological models: Model calibration and "
    "performance assessment. Journal of Hydrology, 600, Article 126674. https://doi.org/10.1016/j.jhydrol.2021.126674",

    "Bhattacharjee, S., Islam, M. T., Kabir, M. E., & Kabir, M. M. (2021). Land-use and land-cover change detection in a "
    "north-eastern wetland ecosystem of Bangladesh using remote sensing and GIS techniques. Earth Systems and "
    "Environment, 5(2), 319\u2013340. https://doi.org/10.1007/s41748-021-00228-3",

    "Datta, S., Karmakar, S., Mezbahuddin, S., Chaudhary, B. S., Hossain, M. M., Hoque, M. E., Abdullah-Al-Mamun, M. M., "
    "& Baul, T. K. (2022). The limits of watershed delineation: Implications of different DEMs, DEM resolutions, and area "
    "threshold values. Hydrology Research, 53(8), 1047\u20131062. https://doi.org/10.2166/nh.2022.126",

    "Hamdan, A. N. A., Almuktar, S., & Scholz, M. (2021). Rainfall-runoff modeling using the HEC-HMS model for the "
    "Al-Adhaim River catchment, northern Iraq. Hydrology, 8(2), Article 58. https://doi.org/10.3390/hydrology8020058",

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

    "Horton, P., Schaefli, B., & Kauzlaric, M. (2022). Why do we have so many different hydrological models? A review "
    "based on the case of Switzerland. WIREs Water, 9(1), Article e1574. https://doi.org/10.1002/wat2.1574",

    "Ichhamati now a trickle. (2019). The Daily Star. "
    "https://www.thedailystar.net/backpage/news/ichhamati-now-trickle-1828042",

    "Jawad, M. (2024). Evaluation of near real-time Global Precipitation Measurement (GPM) precipitation products for "
    "hydrological modelling and flood inundation mapping of sparsely gauged large transboundary basins\u2014A case study "
    "of the Brahmaputra basin. Remote Sensing, 16(10), Article 1756. https://doi.org/10.3390/rs16101756",

    "Knoben, W. J. M., Freer, J. E., & Woods, R. A. (2019). Technical note: Inherent benchmark or not? Comparing "
    "Nash\u2013Sutcliffe and Kling\u2013Gupta efficiency scores. Hydrology and Earth System Sciences, 23(10), "
    "4323\u20134331. https://doi.org/10.5194/hess-23-4323-2019",

    "Moniruzzaman, M., & Mahalder, B. (2026). Assessing SWAT and HEC-HMS model efficiency for watershed management in "
    "the Atrai-Karatoa River Basin, Bangladesh. Evolving Earth, 4, Article 100136. "
    "https://doi.org/10.1016/j.eve.2026.100136",

    "Nujhat, M., Rayhan, M., & Amin, M. K. (2024). Hydrological modelling and its implication in sustainable water "
    "resource management in Gumti River Basin in Bangladesh. International Journal of Sustainability in Energy and "
    "Environment, 1(2), 40\u201348.",

    "Parvez, M., Sadat, N., Tasnim, F., & Nejhum, I. J. (2021). Identifying the causes of waterlogging on people\u2019s "
    "perception towards a resilient community: A case study on Pabna Municipality, Bangladesh. Ecofeminism and Climate "
    "Change, 2(3), 110\u2013126. https://doi.org/10.1108/EFCC-11-2020-0033",

    "Pradhan, R. K., Markonis, Y., Vargas Godoy, M. R., Villalba-Pradas, A., Andreadis, K. M., Nikolopoulos, E. I., "
    "Papalexiou, S. M., Rahim, A., Tapiador, F. J., & Hanel, M. (2022). Review of GPM IMERG performance: A global "
    "perspective. Remote Sensing of Environment, 268, Article 112754. https://doi.org/10.1016/j.rse.2021.112754",

    "Raihan, F., Beaumont, L. J., Maina, J., Saiful Islam, A. K. M., & Harrison, S. P. (2020). Simulating streamflow in "
    "the Upper Halda Basin of southeastern Bangladesh using SWAT model. Hydrological Sciences Journal, 65(1), 138\u2013151. "
    "https://doi.org/10.1080/02626667.2019.1682149",

    "Salimi, S., Almuktar, S. A. A. A. N., & Scholz, M. (2021). Impact of climate change on wetland ecosystems: A "
    "critical review of experimental wetlands. Journal of Environmental Management, 286, Article 112160. "
    "https://doi.org/10.1016/j.jenvman.2021.112160",

    "Shen, H., Tolson, B. A., & Mai, J. (2022). Time to update the split-sample approach in hydrological model "
    "calibration. Water Resources Research, 58(3), Article e2021WR031523. https://doi.org/10.1029/2021WR031523",

    "Soulis, K. X. (2021). Soil Conservation Service Curve Number (SCS-CN) method: Current applications, remaining "
    "challenges, and future perspectives. Water, 13(2), Article 192. https://doi.org/10.3390/w13020192",

    "Tassew, B. G., Belete, M. A., & Miegel, K. (2019). Application of HEC-HMS model for flow simulation in the Lake "
    "Tana Basin: The case of Gilgel Abay catchment, Upper Blue Nile Basin, Ethiopia. Hydrology, 6(1), Article 21. "
    "https://doi.org/10.3390/hydrology6010021",

    "Tk 1,554cr project to revive dying Ichamati. (2024). The Daily Star. "
    "https://www.thedailystar.net/news/bangladesh/news/tk-1554cr-project-revive-dying-ichamati-3555931",

    "U.S. Army Corps of Engineers, Hydrologic Engineering Center. (n.d.). HEC-HMS technical reference manual. Retrieved "
    "August 16, 2026, from https://www.hec.usace.army.mil/confluence/hmsdocs/hmstrm",

    "Uuemaa, E., Ahi, S., Montibeller, B., Muru, M., & Kmoch, A. (2020). Vertical accuracy of freely available global "
    "digital elevation models (ASTER, AW3D30, MERIT, TanDEM-X, SRTM, and NASADEM). Remote Sensing, 12(21), Article 3482. "
    "https://doi.org/10.3390/rs12213482",
]

for r_text in references:
    ref(r_text)

OUT = "thesis/Literature_Review_Gumai_Beel_Ichamoti_HEC-HMS.docx"
doc.save(OUT)

words = sum(len(p.text.split()) for p in doc.paragraphs)
print(f"Saved {OUT}")
print(f"Approximate word count (paragraphs only): {words}")
