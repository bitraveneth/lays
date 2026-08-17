"""Generate the MS-thesis literature review chapter as an APA-7 formatted .docx.

Revision 6: every paragraph carries at least one citation; prose register is
fluent mid-level academic English. Sections 2.1-2.2 are the author's text (with
citations added to previously uncited paragraphs). All cited works were verified
against publisher records (authors, journal, volume, issue, pages, DOI).
"""

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Times New Roman"

doc = Document()

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

header_p = doc.sections[0].header.paragraphs[0]
header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = header_p.add_run()
fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_end)
run.font.name = FONT
run.font.size = Pt(12)


def body(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def equation(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.italic = True
    return p


def h1(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(14)
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


# ================================================================= CHAPTER 2
h1("2. Literature review")

# ------------------------------------------------------------------- 2.1
h2("2.1 Concept of rainfall-runoff modeling")
body(
    "Rainfall\u2013runoff modeling is an important component of hydrological science because it provides a quantitative "
    "framework for understanding how precipitation is transformed into runoff within a watershed. The relationship between "
    "rainfall and runoff is controlled by several factors, including rainfall intensity and duration, soil characteristics, "
    "antecedent moisture conditions, land use and land cover, topography, drainage characteristics, evapotranspiration, and "
    "surface-water storage. Because these factors vary spatially and temporally, the response of a watershed to a given "
    "rainfall event is rarely linear or uniform. Hydrological models are therefore widely used to simplify and represent "
    "these processes and to estimate streamflow from available meteorological and watershed information (Sahu et al., 2023)."
)
body(
    "Rainfall\u2013runoff models have become increasingly important for flood forecasting, watershed management, "
    "water-resources planning and assessment of hydrological hazards. Recent research has produced a wide range of "
    "empirical and conceptual based data-driven models. However, model selection depends strongly on the purpose of the "
    "study, the characteristics of the watershed, the temporal and spatial resolution of available data, and the level of "
    "process representation required. Greater model complexity does not necessarily result in better performance because "
    "model uncertainty can arise from input data, parameterization, structural assumptions, and spatial heterogeneity "
    "(Sahu et al., 2023)."
)
body(
    "Among the available rainfall\u2013runoff models, the Hydrologic Engineering Center\u2013Hydrologic Modeling System "
    "(HEC-HMS) has been widely applied because of its flexible and modular structure. The model allows precipitation to be "
    "converted into runoff through different loss, transform, baseflow, and routing methods. An appropriate representation "
    "can be selected according to the characteristics and objectives of a particular watershed. Recent reviews have "
    "identified the Soil Conservation Service Curve Number (SCS-CN) and SCS Unit Hydrograph methods as among the most "
    "frequently used approaches in HEC-HMS applications, while recession methods and Muskingum-type routing approaches are "
    "also commonly applied (Labade et al., 2025; Sahu et al., 2023; Turkar et al., 2025)."
)
body(
    "The integration of HEC-HMS with Geographic Information Systems (GIS) has further increased its applicability. GIS "
    "enables the extraction of watershed boundaries, drainage networks, slopes, flow paths, land-use information, and soil "
    "characteristics from spatial datasets. These data can subsequently be used to develop hydrological parameters for the "
    "HEC-HMS model. Integrated ArcGIS-derived land-use, soil, and slope information with rainfall and runoff observations "
    "are used to develop HEC-HMS models for the Punpun River Basin (Ranjan & Singh, 2022). Their study demonstrated the "
    "practical value of combining spatial watershed information with observed hydrometeorological data for "
    "rainfall\u2013runoff simulation."
)
body(
    "The application of HEC-HMS is particularly relevant to Bangladesh because of the country\u2019s monsoonal rainfall "
    "regime, extensive river network, low-lying floodplains, wetlands, and recurrent flooding. Several recent studies have "
    "applied HEC-HMS to Bangladesh river basins, including the Khowai, Halda, Old Brahmaputra, Gumti, and "
    "Atrai\u2013Karatoa systems. These studies demonstrate the applicability of the model under Bangladesh\u2019s "
    "hydrological conditions but also highlight challenges related to rainfall representation, discharge availability, "
    "baseflow simulation, and model parameterization (Haque et al., 2024; Moniruzzaman & Mahalder, 2026; Nujhat et al., "
    "2024; Nur et al., 2022; Zhang et al., 2022)."
)
body(
    "The present study focuses on rainfall\u2013runoff modelling of the Gumai Beel\u2013Ichamoti River system using "
    "HEC-HMS. Gumai Beel represents a low-lying wetland and agricultural environment, while the Ichamoti River forms an "
    "important drainage component of the surrounding landscape. The hydrological response of such a system may differ from "
    "conventional upland watersheds because wetland storage, seasonal inundation, drainage connectivity, and agricultural "
    "land use can influence the timing and magnitude of runoff (Adnan et al., 2020; Tang et al., 2020). Therefore, a "
    "review of rainfall\u2013runoff processes, hydrological modelling approaches, HEC-HMS methodology, GIS integration, "
    "wetland hydrology, and previous Bangladesh applications is necessary to establish the scientific basis for the "
    "present research (Labade et al., 2025; Sahu et al., 2023)."
)

# ------------------------------------------------------------------- 2.2
h2("2.2 Rainfall-runoff process")
body(
    "Rainfall\u2013runoff transformation governs the hydrological routing of precipitation across a drainage basin and its "
    "ultimate conversion into streamflow discharge. Incident precipitation across the land surface is partitioned "
    "dynamically among canopy interception, infiltration into the vadose zone, surface depression storage, "
    "evapotranspiration loss, and overland or subsurface lateral flow. The volumetric fraction of precipitation that "
    "reaches the channel network is dictated by the dynamic equilibrium among these constituent processes. Consequently, "
    "the hydrograph configuration produced by a given precipitation event is governed not only by gross rainfall depth, but "
    "also by storm duration, temporal hyetograph distribution, peak rainfall intensity, and the antecedent hydrological "
    "state of the catchment (Labade et al., 2025)."
)
body(
    "Precipitation intensity directly regulates the threshold between infiltration and direct overland flow. When "
    "precipitation rates surpass the saturated hydraulic conductivity and infiltration capacity of the soil, "
    "infiltration-excess (Hortonian) overland flow is initiated and rapidly routed to the channel network. Conversely, "
    "rainfall rates below the infiltration threshold primarily replenish soil moisture storage or contribute to subsurface "
    "matrix flow. Antecedent moisture conditions further modulate this response: a dry catchment exhibits substantial "
    "initial abstraction and retention capacity, whereas antecedent saturation markedly reduces soil matrix suction "
    "potential, accelerating saturation-excess overland flow. As a result, identical storm depths falling under differing "
    "initial moisture conditions generate fundamentally disparate runoff hydrographs (Shi & Wang, 2020)."
)
body(
    "Catchment physiography and surface cover impose critical boundary controls on hydrograph translation and attenuation. "
    "Parameters such as soil texture, effective porosity, hydraulic conductivity, soil profile depth, topographic gradient, "
    "drainage density, and surface roughness collectively determine the velocity and storage dynamics of routed water. "
    "Anthropogenic land-use and land-cover (LULC) modifications directly influence these hydrologic controls such as urban "
    "expansion increases surface imperviousness and accelerates hydrograph. Dense vegetative canopies enhance interception "
    "and infiltration capacity. In agricultural floodplains, hydrological responses exhibit pronounced seasonal variability "
    "driven by cyclical crop growth stages, tillage practices, and artificial field drainage regimes (Dibaba et al., 2020)."
)
body(
    "In low-lying wetland such as Gumai Beel, runoff generation is influenced by dynamic depression storage, delayed "
    "infiltration, and temporary surface-water retention. Within the floodplain morphology of Bangladesh, river\u2013wetland "
    "systems function as hydrodynamic buffers whose storage behavior is dictated by seasonal monsoon inundation pulses and "
    "dynamic lateral river connectivity (Adnan et al., 2020; Akter & Sawon, 2024). These seasonal fluctuations in backwater "
    "staging, storage volume, and hydraulic connectivity introduce significant attenuation and lag into the catchment "
    "discharge hydrograph (Akter & Sawon, 2024). It represents a critical hydrodynamic mechanism in the rainfall\u2013runoff "
    "modeling of the Gumai Beel\u2013Ichamoti River system."
)

# ------------------------------------------------------------------- 2.3
h2("2.3 Factors influencing rainfall\u2013runoff response")

h3("2.3.1 Rainfall characteristics")
body(
    "Rainfall is the primary meteorological input to rainfall\u2013runoff models, and its depth, intensity, duration, "
    "frequency, and temporal and spatial distribution together determine the magnitude and timing of runoff. Small "
    "watersheds under intense rainfall usually produce rapid, sharply peaked responses. In contrast, larger or flatter "
    "watersheds respond more slowly because greater storage and longer travel distances attenuate and delay the flood wave "
    "(Sahu et al., 2023; Souley Tangam et al., 2024)."
)
body(
    "The temporal resolution of the rainfall input must be chosen deliberately. High-resolution records capture "
    "short-duration storm bursts and rainfall peaks, whereas daily totals smooth this variability and may understate the "
    "intensity of convective events. The appropriate resolution therefore depends on the response time of the watershed "
    "and on the purpose of the simulation. Evidence from the Punpun River Basin illustrates this dependence: models built "
    "at daily, monthly, and monsoonal scales from the same data produced clearly different performance, with the monthly "
    "formulation performing best (Ranjan & Singh, 2022)."
)
body(
    "The spatial representation of rainfall requires equal care, because a single rain gauge rarely represents the "
    "precipitation received across a heterogeneous catchment. Areal precipitation is therefore estimated from multiple "
    "stations using Thiessen polygons, inverse-distance weighting, or gridded rainfall products. In the Sirba River Basin "
    "of West Africa, a daily HEC-HMS model driven by observations from 13 meteorological stations confirmed that rainfall "
    "representation is a central component of successful rainfall\u2013runoff modelling (Souley Tangam et al., 2024)."
)
body(
    "Uncertainty in the precipitation input propagates directly into simulated streamflow. Even a well-calibrated model "
    "will produce unreliable output if the forcing does not represent the rainfall actually received by the catchment. "
    "This issue is most serious in data-limited regions where gauges are sparse, a condition that applies to many "
    "Bangladeshi catchments, including the smaller basins for which HEC-HMS has been used to generate runoff information "
    "from limited observations (Nur et al., 2022)."
)

h3("2.3.2 Soil characteristics")
body(
    "Soil properties strongly influence the partitioning of rainfall between infiltration and surface runoff. Soil "
    "texture, hydraulic conductivity, porosity, profile depth, and antecedent moisture together define how much rainfall "
    "the soil column can absorb and store, and therefore how much excess remains for overland flow. Antecedent moisture is "
    "especially influential, and modified curve-number formulations that account for soil moisture and storm duration have "
    "improved runoff prediction substantially compared with the standard method (Shi & Wang, 2020)."
)
body(
    "Within the SCS-CN framework, soil information enters the model through hydrologic soil groups, which rank soils by "
    "infiltration and runoff potential; soils with low infiltration capacity receive higher runoff potential and higher "
    "curve numbers (Soulis, 2021). Soil data are therefore combined routinely with land-use and land-cover information to "
    "derive composite curve numbers for HEC-HMS. In the Punpun River Basin, soil, land-use, and slope layers prepared in "
    "ArcGIS supplied the curve numbers for a set of HEC-HMS models whose performance, judged by R\u00b2, NSE, PBIAS, and "
    "RSR, confirmed the value of spatial soil and land-use information for rainfall\u2013runoff simulation (Ranjan & "
    "Singh, 2022)."
)

h3("2.3.3 Land use and land cover")
body(
    "Land use and land cover shape the rainfall\u2013runoff response by controlling interception, infiltration, surface "
    "roughness, evapotranspiration, and surface storage. Forested surfaces generally intercept and infiltrate more "
    "rainfall than developed surfaces, while agricultural land shows variable behaviour that depends on soil condition, "
    "cultivation practice, crop cover, and seasonal waterlogging (Dibaba et al., 2020)."
)
body(
    "Land-use information is essential for SCS-CN-based modelling because curve numbers are assigned to combinations of "
    "land-use class and hydrologic soil group, so any change in land use redistributes runoff potential across the "
    "catchment (Soulis, 2021). Recent applications have therefore integrated land-use and soil datasets through GIS. In "
    "the Punpun River Basin, LULC, soil, and slope maps prepared in ArcGIS supplied the curve numbers for HEC-HMS (Ranjan "
    "& Singh, 2022), and spatially distributed curve-number estimation from GIS-derived land-use and soil layers has "
    "likewise supported urban flood simulation (Jawale & Thube, 2025)."
)
body(
    "For Gumai Beel, land-use characteristics are particularly relevant because the catchment combines agricultural "
    "activity with seasonal wetland conditions. Land in such environments alternates between cultivation and inundation "
    "across the year, so the effective runoff response of the catchment is itself seasonal, as shown for the beels of the "
    "Ganges\u2013Brahmaputra\u2013Meghna delta, where hydrological behaviour varies strongly with the seasonal flow regime "
    "(Islam et al., 2021). Land-use information therefore provides an essential spatial basis for parameterising the "
    "rainfall\u2013runoff model."
)

h3("2.3.4 Topography")
body(
    "Topography controls the direction, velocity, and concentration of water movement within a watershed. Elevation, "
    "slope, drainage density, stream order, flow length, and basin shape together determine the time required for runoff "
    "to reach the outlet, and hence the timing and shape of the outlet hydrograph (Sahu et al., 2023)."
)
body(
    "Digital Elevation Models (DEMs) are therefore central to GIS-based hydrological modelling. Processing a DEM yields "
    "flow direction, flow accumulation, stream networks, watershed boundaries, sub-basins, and longest flow paths, from "
    "which the HEC-HMS basin model is constructed. DEM-derived flow direction, flow accumulation, stream ordering, and "
    "basin delineation form the foundation of current GIS\u2013HEC-HMS workflows (Jawale & Thube, 2025). Delineation "
    "outcomes, however, remain sensitive to the choice of DEM product, its resolution, and the stream-definition "
    "threshold, and this sensitivity is greatest in low-relief terrain (Datta et al., 2022)."
)
body(
    "Topography matters most in low-gradient wetland systems, where elevation differences of only a few decimetres can "
    "decide where water accumulates and how it drains. Because DEM error in such terrain can approach the topographic "
    "signal itself, accurate DEM processing and careful delineation of the Gumai Beel catchment are essential components "
    "of the present study (Datta et al., 2022)."
)

# ------------------------------------------------------------------- 2.4
h2("2.4 Hydrological modelling")
body(
    "Hydrological models provide simplified representations of the natural hydrological cycle and are used to estimate "
    "runoff, streamflow, soil moisture, evapotranspiration, groundwater flow, and storage. Models differ both in how they "
    "represent hydrological processes and in their spatial and temporal structure (Sahu et al., 2023)."
)
body(
    "With respect to process representation, models are broadly classified as empirical, conceptual, physically based, or "
    "data-driven. Empirical models establish statistical relationships between observed inputs and outputs; conceptual "
    "models represent the major processes through simplified storage structures; and physically based models solve "
    "governing physical equations with spatially distributed parameters. Artificial intelligence and machine-learning "
    "approaches have more recently been used to capture nonlinear relationships between meteorological inputs and "
    "streamflow. Spatially, models are categorised as lumped, semi-distributed, or distributed, according to whether the "
    "watershed is treated as one homogeneous unit, divided into sub-units, or resolved over an explicit grid (Sahu et "
    "al., 2023)."
)
body(
    "Model complexity should match the characteristics of the watershed and the objective of the study. HEC-HMS has been "
    "identified as a practical choice for dendritic drainage systems, with the SCS-CN and Soil Moisture Accounting methods "
    "dominating event-based and continuous applications respectively (Sahu et al., 2023). Model choice must also reflect "
    "data availability, because heavily parameterised models do not necessarily outperform simpler conceptual structures "
    "when observations are limited (Odey & Cho, 2025). This consideration is directly relevant to Bangladesh, where "
    "discharge measurements are spatially sparse and many local catchments remain poorly gauged (Nur et al., 2022)."
)

# ------------------------------------------------------------------- 2.5
h2("2.5 HEC-HMS rainfall\u2013runoff model")
body(
    "HEC-HMS is a watershed-scale hydrological modelling system developed to simulate precipitation\u2013runoff "
    "processes. The watershed is represented as a network of interconnected hydrological elements, including sub-basins, "
    "reaches, junctions, reservoirs, sources, and sinks (U.S. Army Corps of Engineers [USACE], n.d.). The principal "
    "advantage of the system is its modular structure: alternative methods can be selected for precipitation losses, "
    "runoff transformation, baseflow, and channel routing, which allows the model to be adapted to different watershed "
    "types and objectives (Sahu et al., 2023)."
)
body(
    "The model has been applied successfully under diverse climatic and physiographic conditions and performs "
    "particularly well in dendritic drainage systems, provided that an appropriate loss method is selected (Sahu et al., "
    "2023). Across the application literature, the SCS-CN and SCS Unit Hydrograph methods dominate loss estimation and "
    "runoff transformation, Muskingum and recession approaches are most common for routing and baseflow, and integration "
    "with GIS, remote sensing, and parameter optimisation continues to expand (Labade et al., 2025). Reliable performance "
    "under varied hydrological conditions has been confirmed repeatedly, with calibration quality, method selection, and "
    "the integration of hydro-meteorological and geospatial datasets identified as the decisive factors (Turkar et al., "
    "2025). These characteristics make HEC-HMS suitable for the present study, in which a rainfall\u2013runoff model is "
    "developed from rainfall and discharge observations together with spatial watershed information."
)

# ------------------------------------------------------------------- 2.6
h2("2.6 HEC-HMS model components")

h3("2.6.1 Basin model")
body(
    "The basin model provides the spatial and hydrological structure of the watershed. It represents the sub-basins and "
    "drainage network through which runoff is generated and conveyed, and its configuration depends on watershed "
    "morphology, drainage structure, the location of observation points, and the modelling objective (USACE, n.d.). For "
    "the Gumai Beel\u2013Ichamoti system, delineation is especially important because the basin model must capture the "
    "main drainage pathways connecting the beel to the river. DEM-based GIS processing provides the flow accumulation, "
    "stream networks, sub-basin boundaries, and outlet locations from which this structure is built (Jawale & Thube, "
    "2025)."
)

h3("2.6.2 Meteorological model")
body(
    "The meteorological model supplies precipitation and other meteorological inputs to HEC-HMS, using either gauge "
    "observations or gridded products according to availability (USACE, n.d.). Spatial representation becomes critical "
    "when several rainfall stations fall within the watershed, and Thiessen polygons, gridded rainfall, or other "
    "interpolation methods are then used to estimate mean areal precipitation. A multi-station configuration of this kind "
    "supported the daily simulation of the Sirba River Basin, where rainfall from 13 stations and discharge at one outlet "
    "station over 2006\u20132020 enabled a comparison of continuous and event-based simulation (Souley Tangam et al., "
    "2024)."
)

h3("2.6.3 Loss model")
body(
    "The loss model estimates the fraction of rainfall that does not become surface runoff, accounting for infiltration, "
    "interception, and initial abstraction. Available methods include the SCS Curve Number, Initial and Constant, "
    "Green\u2013Ampt, Deficit and Constant, and Soil Moisture Accounting formulations, and the choice among them depends "
    "on the modelling objective and the information available (USACE, n.d.)."
)
body(
    "The SCS-CN method is the most widely used loss method in HEC-HMS because it provides a simple, well-documented "
    "relationship between rainfall excess and watershed characteristics; successive reviews confirm its dominance in "
    "event-based modelling (Labade et al., 2025; Sahu et al., 2023). The curve number combines the influence of soil, "
    "land use, land treatment, and antecedent moisture on runoff potential: higher values indicate greater runoff, and "
    "lower values indicate greater infiltration and storage (Soulis, 2021). The method is relevant to the present study "
    "because the Gumai Beel catchment contains agricultural and wetland land uses whose soil and land-use characteristics "
    "can be mapped spatially through GIS (Ranjan & Singh, 2022)."
)

h3("2.6.4 Transform method")
body(
    "The transform method converts excess rainfall into direct runoff and thereby determines the shape and timing of the "
    "simulated hydrograph. The SCS Unit Hydrograph is the most frequently applied transform in HEC-HMS: it requires only "
    "the basin lag time and applies a standardised dimensionless relationship to distribute excess precipitation in time "
    "(Labade et al., 2025). Alternatives include the Clark and Snyder unit hydrographs, ModClark, and kinematic-wave "
    "approaches, and the choice depends on watershed characteristics and available data (USACE, n.d.). The combination of "
    "the SCS Unit Hydrograph with SCS-CN losses and recession baseflow remains the standard configuration in "
    "GIS-supported rainfall\u2013runoff and flood modelling (Jawale & Thube, 2025)."
)

h3("2.6.5 Baseflow")
body(
    "Baseflow is the sustained component of streamflow that is not produced directly by the current rainfall event. It "
    "originates from groundwater discharge, delayed subsurface flow, and other forms of watershed storage, and it is "
    "difficult to represent because groundwater processes are usually poorly observed (Odey & Cho, 2025). Discrepancies "
    "between simulated and observed hydrographs therefore concentrate in the recession limb. This difficulty is evident "
    "in the Halda River catchment, where an HEC-HMS model agreed satisfactorily with observed discharge on several "
    "indicators yet reproduced the baseflow component poorly during calibration (Haque et al., 2024). The issue is "
    "potentially important for the Gumai Beel\u2013Ichamoti system, where wetland storage and subsurface interaction may "
    "sustain flow between rainfall events (Islam et al., 2021)."
)

h3("2.6.6 Channel routing")
body(
    "Channel routing describes the movement of runoff through the river network and accounts for the translation and "
    "attenuation of the hydrograph caused by channel storage and travel time. The Muskingum method is used most "
    "frequently in HEC-HMS because it represents these effects through only two parameters, the travel-time and storage "
    "parameter K and the weighting factor X (Labade et al., 2025). In Bangladesh, Muskingum routing combined with SCS-CN "
    "losses supported the calibrated and validated simulation of the Gumti River Basin from observed rainfall and "
    "discharge records (Nujhat et al., 2024). Routing is important for the present study because the objective includes "
    "not only runoff generation within the catchment but also the movement of that flow through the Ichamoti drainage "
    "network (Zhang et al., 2022)."
)

# ------------------------------------------------------------------- 2.7
h2("2.7 SCS Curve Number method in rainfall\u2013runoff modelling")
body(
    "The SCS Curve Number method is an empirical approach for estimating direct runoff from rainfall through a single "
    "parameter, the curve number, which represents watershed characteristics. Its modest data requirements and its "
    "compatibility with GIS-based parameterisation explain its continued dominance in applied practice (Soulis, 2021). "
    "Runoff response is controlled by the potential maximum retention S, which is inversely related to the curve number; "
    "the curve number itself is determined from land use, hydrologic soil group, land treatment, and antecedent moisture "
    "condition (Soulis, 2021). The governing relationship is expressed as"
)
equation("Q = (P \u2212 Ia)\u00b2 / [(P \u2212 Ia) + S]")
body(
    "where Q is the direct runoff depth, P is the precipitation depth, Ia is the initial abstraction, and S is the "
    "potential maximum retention. In the conventional formulation, the initial abstraction is taken as a fixed fraction "
    "of the potential retention, and the HEC-HMS implementation allows both the curve number and the initial abstraction "
    "to be specified explicitly (USACE, n.d.)."
)
body(
    "The appeal of the method for GIS-based modelling lies in its ability to translate spatial soil and land-use "
    "information directly into runoff potential. ArcGIS-derived soil and LULC layers supplied the curve numbers for a "
    "successful HEC-HMS application in the Punpun River Basin (Ranjan & Singh, 2022), and the combination of SCS-CN with "
    "the SCS Unit Hydrograph and recession baseflow continues to support flood-oriented modelling (Jawale & Thube, 2025). "
    "The method nevertheless has recognised limitations. It compresses complex infiltration and storage processes into a "
    "single parameter, the conventional initial abstraction ratio of 0.2 has been found in many settings to overestimate "
    "initial losses, and the method does not track the evolution of soil moisture through long continuous simulations "
    "(Shi & Wang, 2020; Soulis, 2021). It should therefore be applied with attention to the temporal scale of the "
    "simulation and the characteristics of the study area."
)

# ------------------------------------------------------------------- 2.8
h2("2.8 GIS integration in HEC-HMS modelling")
body(
    "GIS has become an integral component of hydrological modelling because watershed characteristics are inherently "
    "spatial. It provides the tools for converting DEMs, land-use maps, soil maps, and drainage networks into the "
    "parameters that a hydrological model requires. A typical GIS-supported HEC-HMS workflow begins with DEM "
    "preprocessing, in which sink filling and flow-direction analysis are followed by flow-accumulation analysis and "
    "stream-network extraction; the resulting network is then used to delineate sub-basins and locate the watershed "
    "outlet (Jawale & Thube, 2025)."
)
body(
    "The conversion of spatial information into model parameters is well established in recent practice. ArcGIS-based "
    "preparation of LULC, soil, and slope layers supplied the curve numbers for HEC-HMS simulation of the Punpun River "
    "Basin (Ranjan & Singh, 2022), and DEM-derived delineation combined with soil and land-use overlay produced spatially "
    "distributed curve numbers for urban flood modelling (Jawale & Thube, 2025). The reliability of this workflow, "
    "however, rests on the terrain analysis at its head, because delineation outcomes in low-relief terrain are sensitive "
    "to DEM choice, resolution, and threshold settings (Datta et al., 2022). GIS integration is therefore central to the "
    "present study, in which the Gumai Beel catchment must be delineated from topographic data and its spatial "
    "characteristics incorporated into the rainfall\u2013runoff model."
)

# ------------------------------------------------------------------- 2.9
h2("2.9 Event-based and continuous HEC-HMS simulation")
body(
    "HEC-HMS supports both event-based and continuous rainfall\u2013runoff simulation. Event-based modelling reproduces "
    "the response to individual storms and is standard for flood estimation and design-storm analysis, whereas continuous "
    "modelling simulates watershed behaviour over long periods and therefore requires representation of changing soil "
    "moisture, groundwater, and evapotranspiration. A review of applications from 2000 to 2023 confirms that event-based "
    "simulations rely mainly on the SCS-CN and SCS Unit Hydrograph methods, that continuous simulations require dynamic "
    "soil-moisture and baseflow representations such as Soil Moisture Accounting and linear-reservoir baseflow, and that "
    "the choice between modes should follow from the intended application and the available data (Odey & Cho, 2025)."
)
body(
    "A direct empirical comparison of the two modes is available for the Sirba River Basin, where a continuous simulation "
    "over 2006\u20132020 was evaluated alongside event-based simulations of selected major floods. Both schemes performed "
    "satisfactorily, but the event-based simulations were more accurate, with R\u00b2 of 0.94\u20130.98 against "
    "0.84\u20130.87 for the continuous case, and the calibrated parameter distributions differed between schemes, showing "
    "that the modelling mode itself conditions parameter estimation (Souley Tangam et al., 2024)."
)
body(
    "For the Gumai Beel\u2013Ichamoti system, the appropriate mode follows from the temporal resolution and duration of "
    "the available rainfall and discharge records (Odey & Cho, 2025). Event-based simulation is suitable if the main "
    "purpose is to reproduce individual high-flow events, whereas sufficiently long continuous observations would permit "
    "a continuous simulation that captures the full seasonal cycle of wetland storage and release."
)

# ------------------------------------------------------------------ 2.10
h2("2.10 Calibration and validation of HEC-HMS")
body(
    "Calibration is essential in rainfall\u2013runoff modelling because several HEC-HMS parameters cannot be measured "
    "directly and must be estimated from the observed response of the watershed (Turkar et al., 2025). During "
    "calibration, parameters are adjusted to reduce the difference between observed and simulated discharge. A sound "
    "calibration considers several attributes of the hydrograph at once, including peak discharge, runoff volume, time to "
    "peak, the rising and recession limbs, and low-flow behaviour, because agreement in one statistic can hide "
    "compensating errors elsewhere. Validation against an independent period of observations then tests whether the "
    "calibrated model can reproduce hydrological behaviour under conditions not used for parameter estimation (Sahu et "
    "al., 2023)."
)
body(
    "Model performance is commonly quantified using the Nash\u2013Sutcliffe Efficiency (NSE), the coefficient of "
    "determination (R\u00b2), the root-mean-square error (RMSE), and percent bias (PBIAS), which provide complementary "
    "information on correlation, volume error, and overall fit (Labade et al., 2025). In the Punpun River Basin, "
    "evaluation of daily, monthly, and monsoonal HEC-HMS models against R\u00b2, NSE, PBIAS, and RSR showed generally "
    "satisfactory performance, with the monthly model the strongest (Ranjan & Singh, 2022). In Bangladesh, the "
    "hydrological component of a coupled HEC-HMS\u2013HEC-RAS flood-risk framework for the Old Brahmaputra floodplain "
    "achieved NSE values of 0.93 and 0.81, R\u00b2 values of 0.95 and 0.89, and PBIAS of \u22121.17% and 2.40% for "
    "calibration and validation respectively, confirming that HEC-HMS can supply reliable hydrological input to "
    "flood-risk assessment when adequate observations are available (Zhang et al., 2022)."
)
body(
    "Numerical thresholds alone, however, do not establish model adequacy. Visual comparison of observed and simulated "
    "hydrographs remains necessary, because a model can reach a satisfactory NSE while still misrepresenting particular "
    "flood peaks or recession behaviour that matter for the intended application (Odey & Cho, 2025; Turkar et al., 2025)."
)

# ------------------------------------------------------------------ 2.11
h2("2.11 Parameter sensitivity and model uncertainty")
body(
    "Model parameters differ in their influence on simulated streamflow, and sensitivity analysis identifies the "
    "parameters that control model output so that calibration effort can be directed efficiently. In HEC-HMS, the "
    "influential parameters relate to rainfall losses, basin response, baseflow, and routing. For SCS-CN-based models the "
    "curve number is usually dominant because it directly governs the conversion of rainfall into excess; lag time "
    "controls hydrograph timing and shape; and routing parameters control downstream translation and attenuation. Curve "
    "number, infiltration-related parameters, lag time, and baseflow parameters have been identified consistently as the "
    "components requiring the most careful calibration (Turkar et al., 2025), and sensitivity analysis in the Sirba River "
    "Basin likewise ranked curve number, initial abstraction, lag time, and routing time factors as the most influential "
    "(Souley Tangam et al., 2024)."
)
body(
    "Parameter uncertainty becomes especially important when observations are limited, because different parameter "
    "combinations can produce nearly identical hydrographs, and a satisfactory calibration therefore does not guarantee "
    "that each parameter carries a unique physical interpretation (Sahu et al., 2023). For the Gumai Beel\u2013Ichamoti "
    "system, sensitivity analysis will identify the parameters most influential on simulated discharge and ensure that "
    "calibration is concentrated on the parameters that matter most (Turkar et al., 2025)."
)

# ------------------------------------------------------------------ 2.12
h2("2.12 Recent international applications of HEC-HMS")
body(
    "Applications published since 2020 confirm that HEC-HMS remains in wide use across diverse climatic and physiographic "
    "settings. In the Punpun River Basin of eastern India, daily rainfall and runoff observations combined with "
    "GIS-derived LULC, soil, and slope data supported daily, monthly, and monsoonal models, of which the monthly model "
    "performed best, underlining the influence of temporal scale on simulation quality (Ranjan & Singh, 2022). In the "
    "Sirba River Basin of West Africa, continuous and event-based simulations both achieved satisfactory performance, "
    "with the event-based scheme superior for the selected flood events (Souley Tangam et al., 2024). In a Romanian "
    "catchment with increased groundwater discharge potential, the SCS-CN, SCS Unit Hydrograph, and Muskingum "
    "configuration proved applicable even where groundwater processes influence the catchment response (Herbei et al., "
    "2024)."
)
body(
    "Integration with hydraulic modelling has extended these applications further. In the Cypress Creek watershed of "
    "Texas, wetlands represented as reservoirs within HEC-HMS were combined with HEC-RAS river hydraulics to examine the "
    "effect of wetland size and location on watershed-scale flood control; upstream placement and greater storage reduced "
    "downstream flood area, depth, and duration (Tang et al., 2020). This finding is conceptually important for "
    "wetland-dominated catchments because it shows quantitatively that wetland storage modifies the downstream "
    "hydrological response. Taken together, these studies establish that HEC-HMS is not restricted to conventional river "
    "catchments: its flexible representation of sub-basins, storage, losses, transformation, and routing makes it "
    "applicable to systems in which wetlands regulate runoff behaviour (Odey & Cho, 2025; Tang et al., 2020)."
)

# ------------------------------------------------------------------ 2.13
h2("2.13 Wetland and beel hydrology")
body(
    "Wetlands are important components of watershed hydrology because they store precipitation and runoff temporarily and "
    "thereby influence the timing and magnitude of downstream flows. The strength of this regulation depends on wetland "
    "size, location, storage capacity, connectivity, vegetation, soil characteristics, and the relationship with "
    "surrounding rivers (Tang et al., 2020). Hydrologic\u2013hydraulic simulation of the Cypress Creek watershed "
    "demonstrates the principle directly: larger wetlands, and wetlands located farther upstream, reduced downstream "
    "inundation extent, depth, and duration (Tang et al., 2020). Although derived from a North American case, the result "
    "expresses a general mechanism, namely that wetland storage alters the timing and magnitude of runoff delivered to "
    "downstream channels."
)
body(
    "Wetlands are particularly significant in Bangladesh, whose floodplains contain numerous seasonal and permanent "
    "depressional water bodies. Beels are saucer-shaped depressions that retain water during the wet season and dry "
    "partially or completely during the dry season, and their behaviour is closely linked to seasonal flooding and river "
    "connectivity (Adnan et al., 2020). Simulation of beels within the polders of the Ganges\u2013Brahmaputra\u2013Meghna "
    "delta has revealed strong spatial and seasonal variability in their hydrological and sedimentation behaviour, "
    "governed by the seasonal flow regime and the interaction between the beel and the surrounding river system (Islam et "
    "al., 2021)."
)
body(
    "These characteristics distinguish beel systems from conventional upland watersheds. In a beel-dominated catchment, "
    "rainfall typically contributes first to temporary surface storage and only later drains to downstream channels, so "
    "the runoff peak is delayed and attenuated relative to an equivalent upland response (Islam et al., 2021; Tang et "
    "al., 2020). For rainfall\u2013runoff modelling this creates a specific challenge: a conventional model may reproduce "
    "runoff generation satisfactorily yet require careful representation of storage and routing to reproduce downstream "
    "hydrograph timing. Understanding the hydrological character of the wetland is therefore a prerequisite for selecting "
    "the appropriate HEC-HMS structure (Odey & Cho, 2025)."
)

# ------------------------------------------------------------------ 2.14
h2("2.14 Hydrological modelling in Bangladesh")
body(
    "Bangladesh has a complex hydrological environment characterised by monsoonal rainfall, extensive floodplains, a "
    "dense river network, abundant wetlands, and strong seasonal variation in water levels and discharge. Flooding recurs "
    "annually and is a major concern for agriculture, infrastructure, ecosystems, and communities (Akter & Sawon, 2024; "
    "Zhang et al., 2022). Reliable rainfall\u2013runoff models are therefore needed to understand watershed behaviour and "
    "support flood management, yet their development is complicated by the limited spatial coverage of rainfall and "
    "discharge observations (Nur et al., 2022)."
)
body(
    "Recent studies document the expanding application of HEC-HMS across the country. For the flash-flood-prone Khowai "
    "River Basin, a calibrated and validated HEC-HMS model generated runoff information from available rainfall data "
    "under conditions of limited discharge observation (Nur et al., 2022). For the Gumti River Basin, SRTM-based "
    "delineation, SCS-CN losses, and Muskingum routing supported calibration against 2019\u20132020 observations and "
    "validation against 2021 records, with performance evaluated through R\u00b2, NSE, RSR, and PBIAS and the calibrated "
    "model recommended for flood prediction and water-resources planning (Nujhat et al., 2024). In the Halda River "
    "catchment, a cascade-reservoir HEC-HMS framework with curve numbers optimised against SWAT results achieved "
    "agreement with observed discharge across several indicators, although baseflow representation remained a documented "
    "weakness (Haque et al., 2024). Hydrological models of selected flash-flood-prone rivers have further advanced the "
    "understanding of rapid-response behaviour under monsoonal forcing (Akter & Sawon, 2024). Together, these studies "
    "confirm that HEC-HMS can be applied successfully in Bangladesh while showing that performance depends on careful "
    "calibration and appropriate representation of watershed processes."
)

# ------------------------------------------------------------------ 2.15
h2("2.15 HEC-HMS and flood modelling in Bangladesh")
body(
    "The application of HEC-HMS in Bangladesh has expanded beyond rainfall\u2013runoff simulation toward integrated flood "
    "modelling. For the Old Brahmaputra River floodplain, an integrated flood-risk framework coupled HEC-HMS with HEC-RAS "
    "1D/2D hydraulics and a bottom-up vulnerability analysis: the hydrological model generated the flows used in the "
    "hydraulic simulation, and the combined framework produced strong calibration and validation results, demonstrating "
    "the value of integrating hydrological and hydraulic models for flood-risk assessment (Zhang et al., 2022)."
)
body(
    "This form of integration is directly relevant to the present study, because a calibrated rainfall\u2013runoff model "
    "is the foundation for any future hydraulic modelling of the Gumai Beel\u2013Ichamoti system. Once rainfall-generated "
    "discharge can be simulated reliably, the resulting hydrographs can serve as boundary conditions for a hydraulic "
    "model such as HEC-RAS (Zhang et al., 2022). The growing use of HEC-HMS in Bangladesh also reflects a broader shift "
    "toward spatially integrated modelling, in which GIS, DEMs, gridded rainfall products, and remote sensing are "
    "combined with observed discharge to improve model representation (Labade et al., 2025)."
)

# ------------------------------------------------------------------ 2.16
h2("2.16 Comparison of HEC-HMS with SWAT and other models")
body(
    "HEC-HMS is not the only framework available for rainfall\u2013runoff simulation. SWAT, a widely used "
    "semi-distributed watershed model, is particularly suited to long-term assessments involving land-use management, "
    "agricultural practice, sediment transport, and nutrient dynamics. The essential difference lies in their objectives: "
    "HEC-HMS concentrates on precipitation\u2013runoff transformation and hydrograph simulation, whereas SWAT represents "
    "the broader watershed water balance and land-management processes. A direct comparison in the Atrai\u2013Karatoa "
    "River Basin of Bangladesh found that SWAT performed better for high flows under daily simulation while HEC-HMS was "
    "more accurate for medium flows, leading to the conclusion that model selection should follow from the intended "
    "application and the hydrological character of the watershed (Moniruzzaman & Mahalder, 2026)."
)
body(
    "Machine-learning approaches, including artificial neural networks and long short-term memory networks, have also "
    "become popular because they capture nonlinear rainfall\u2013runoff relationships and can achieve high predictive "
    "accuracy where large training datasets exist; their physical interpretability, however, is lower than that of "
    "conceptual hydrological models (Sahu et al., 2023). HEC-HMS is appropriate for the present study because the primary "
    "objective is to represent the rainfall\u2013runoff response of a specific watershed through physically interpretable "
    "components, allowing losses, transformation, baseflow, and routing to be examined separately (Turkar et al., 2025)."
)

# ------------------------------------------------------------------ 2.17
h2("2.17 Challenges of rainfall\u2013runoff modelling in Bangladesh")
body(
    "Four challenges condition rainfall\u2013runoff modelling in Bangladesh. The first is the spatial variability of "
    "rainfall: monsoon precipitation, particularly during convective events, can vary substantially over short distances, "
    "so a limited gauge network introduces uncertainty into estimates of mean areal precipitation (Souley Tangam et al., "
    "2024). The second is the scarcity of discharge data, because many local rivers and wetland systems lack long-term "
    "gauging stations, which complicates calibration and validation and applies with particular force to smaller "
    "catchments such as the Gumai Beel\u2013Ichamoti system (Nur et al., 2022). The third is the complex interaction "
    "among rivers, floodplains, wetlands, and groundwater, which conventional models simplify at the cost of potential "
    "errors in baseflow, storage, and recession behaviour (Haque et al., 2024). The fourth is the seasonal variation of "
    "land use and wetland storage in agricultural wetland environments, where land is cultivated in one season and "
    "inundated in another, so runoff characteristics do not remain constant through the year (Islam et al., 2021)."
)
body(
    "The Bangladeshi studies reviewed above show that HEC-HMS delivers satisfactory rainfall\u2013runoff simulation where "
    "suitable rainfall and discharge records exist, while underscoring the need for careful calibration, appropriate "
    "selection of model components, and explicit attention to local hydrological characteristics (Haque et al., 2024; "
    "Moniruzzaman & Mahalder, 2026; Nujhat et al., 2024; Nur et al., 2022)."
)

# ------------------------------------------------------------------ 2.18
h2("2.18 Research gap")
body(
    "The reviewed literature establishes that HEC-HMS has been applied extensively to rainfall\u2013runoff simulation and "
    "flood assessment under diverse environmental conditions, and that recent research increasingly integrates the model "
    "with GIS, remote sensing, DEM-derived watershed characteristics, and hydraulic models (Labade et al., 2025; Odey & "
    "Cho, 2025). Several gaps nevertheless remain relevant to the present study."
)
body(
    "First, recent HEC-HMS research in Bangladesh has concentrated on comparatively large or well-studied basins, "
    "including the Khowai, Halda, Old Brahmaputra, Gumti, and Atrai\u2013Karatoa systems (Haque et al., 2024; "
    "Moniruzzaman & Mahalder, 2026; Nujhat et al., 2024; Nur et al., 2022; Zhang et al., 2022). These studies provide "
    "valuable methodological precedents but do not necessarily represent the behaviour of smaller wetland-dominated "
    "catchments. Second, rainfall\u2013runoff modelling of beel\u2013river systems has received limited attention. "
    "Wetlands and beels act as temporary storage that alters the timing and magnitude of downstream runoff; international "
    "evidence confirms the influence of wetland storage on flood response (Tang et al., 2020), and Bangladeshi evidence "
    "confirms the strong seasonal and spatial variability of beel systems (Islam et al., 2021), yet few studies have "
    "carried these characteristics into a dedicated HEC-HMS framework for a Bangladeshi beel."
)
body(
    "Third, the literature search undertaken for this study identified very limited recent peer-reviewed research "
    "addressing the Gumai Beel\u2013Ichamoti River system with HEC-HMS, indicating a local-scale research gap rather than "
    "a general absence of hydrological modelling research in Bangladesh (Moniruzzaman & Mahalder, 2026; Nujhat et al., "
    "2024). Fourth, although previous Bangladeshi studies demonstrate the value of integrating observed rainfall, "
    "discharge, and GIS-derived DEM, land-use, and soil information, these approaches have been applied mainly to larger "
    "conventional river basins (Ranjan & Singh, 2022; Zhang et al., 2022); their extension to an agricultural wetland "
    "system such as Gumai Beel can therefore yield new information on the rainfall\u2013runoff response of a low-lying "
    "catchment. Finally, a calibrated and validated HEC-HMS model of Gumai Beel would provide the foundation for "
    "subsequent hydrological and hydraulic investigation, because the validated model could be linked with hydraulic "
    "models to examine flood extent, water levels, drainage behaviour, or future scenarios (Zhang et al., 2022)."
)

# ------------------------------------------------------------------ 2.19
h2("2.19 Relevance of the present study")
body(
    "The present study addresses the identified gap by developing a rainfall\u2013runoff model of the Gumai "
    "Beel\u2013Ichamoti River system in HEC-HMS, integrating rainfall and discharge observations with spatial information "
    "derived from GIS and DEM datasets, following the workflow established in recent GIS-supported applications (Jawale & "
    "Thube, 2025; Ranjan & Singh, 2022). Watershed delineation supplies the spatial structure of the model, rainfall "
    "provides the primary meteorological forcing, and observed discharge supports calibration and validation, so that "
    "simulated runoff can be evaluated against appropriate statistical performance indicators (Labade et al., 2025)."
)
body(
    "The study derives particular significance from its setting. Gumai Beel is a low-lying agricultural and wetland "
    "environment whose hydrological response is expected to reflect not only rainfall but also topography, land use, soil "
    "characteristics, drainage connectivity, and temporary wetland storage (Islam et al., 2021; Tang et al., 2020). "
    "HEC-HMS provides an interpretable framework for examining these controls, because rainfall losses, direct-runoff "
    "transformation, baseflow, and routing are represented as separable components whose relative influence can be "
    "investigated individually (Turkar et al., 2025; USACE, n.d.)."
)
body(
    "The resulting model also establishes a foundation for future research. The simulated discharge hydrographs could "
    "serve as inputs to HEC-RAS for hydraulic flood modelling, following the integrated HEC-HMS\u2013HEC-RAS framework "
    "applied to the Old Brahmaputra floodplain (Zhang et al., 2022). The present study therefore contributes both to the "
    "hydrological understanding of Gumai Beel and to the wider application of GIS-supported HEC-HMS modelling in small, "
    "wetland-dominated catchments of Bangladesh (Odey & Cho, 2025; Tang et al., 2020)."
)

# ================================================================ REFERENCES
page_break()
h1("References")

references = [
    "Adnan, M. S. G., Talchabhadel, R., Nakagawa, H., & Hall, J. W. (2020). The potential of Tidal River Management for "
    "flood alleviation in South Western Bangladesh. Science of the Total Environment, 731, Article 138747. "
    "https://doi.org/10.1016/j.scitotenv.2020.138747",

    "Akter, A., & Sawon, F. S. (2024). Hydrological modeling of the selected flash flood-prone rivers. Natural Hazards. "
    "https://doi.org/10.1007/s11069-024-06928-z",

    "Datta, S., Karmakar, S., Mezbahuddin, S., Chaudhary, B. S., Hossain, M. M., Hoque, M. E., Abdullah-Al-Mamun, M. M., "
    "& Baul, T. K. (2022). The limits of watershed delineation: Implications of different DEMs, DEM resolutions, and "
    "area threshold values. Hydrology Research, 53(8), 1047\u20131062. https://doi.org/10.2166/nh.2022.126",

    "Dibaba, W. T., Demissie, T. A., & Miegel, K. (2020). Watershed hydrological response to combined land use/land cover "
    "and climate change in highland Ethiopia: Finchaa catchment. Water, 12(6), Article 1801. "
    "https://doi.org/10.3390/w12061801",

    "Haque, M. B., Karmakar, S., & Hossain, M. M. (2024). Rainfall-runoff modeling using the HEC-HMS flow modeling "
    "framework for the Halda River catchment, Bangladesh [Preprint]. Research Square. "
    "https://doi.org/10.21203/rs.3.rs-3824469/v1",

    "Herbei, M. V., B\u0103d\u0103lu\u021b\u0103-Minda, C., Popescu, C. A., Horablaga, A., Dragomir, L. O., Popescu, G., "
    "Kader, S., & Sestras, P. (2024). Rainfall-runoff modeling based on HEC-HMS model: A case study in an area with "
    "increased groundwater discharge potential. Frontiers in Water, 6, Article 1474990. "
    "https://doi.org/10.3389/frwa.2024.1474990",

    "Islam, M. F., Middelkoop, H., Schot, P. P., Dekker, S. C., & Griffioen, J. (2021). Spatial and seasonal variability "
    "of sediment accumulation potential through controlled flooding of the beels located in the polders of the "
    "Ganges-Brahmaputra-Meghna delta of Southwest Bangladesh. Hydrological Processes, 35(4), Article e14119. "
    "https://doi.org/10.1002/hyp.14119",

    "Jawale, P. S., & Thube, A. D. (2025). Rainfall-runoff modeling of urban floods using GIS and HEC-HMS. MethodsX, 15, "
    "Article 103437. https://doi.org/10.1016/j.mex.2025.103437",

    "Labade, P., Ayare, B. L., Bhange, H. N., Ingle, P. M., & Kolhe, P. R. (2025). A comprehensive review on "
    "rainfall-runoff modelling using HEC-HMS. International Journal of Research in Agronomy, 8(10), 1130\u20131138. "
    "https://doi.org/10.33545/2618060X.2025.v8.i10o.4374",

    "Moniruzzaman, M., & Mahalder, B. (2026). Assessing SWAT and HEC-HMS model efficiency for watershed management in "
    "the Atrai-Karatoa River Basin, Bangladesh. Evolving Earth, 4, Article 100136. "
    "https://doi.org/10.1016/j.eve.2026.100136",

    "Nujhat, M., Rayhan, M., & Amin, M. K. (2024). Hydrological modelling and its implication in sustainable water "
    "resource management in Gumti River Basin in Bangladesh. International Journal of Sustainability in Energy and "
    "Environment, 1(2), 40\u201348.",

    "Nur, F., Mohib, K. M., Toma, U. T., Rozario, P. M., Bari, M. S., Anjum, N., & Khandakar, F. (2022, February "
    "10\u201312). Rainfall-runoff simulation of Khowai River basin using HEC-HMS model [Conference presentation]. 6th "
    "International Conference on Civil Engineering for Sustainable Development (ICCESD 2022), Khulna University of "
    "Engineering & Technology, Khulna, Bangladesh.",

    "Odey, G., & Cho, Y. (2025). Event-based vs. continuous hydrological modeling with HEC-HMS: A review of use cases, "
    "methodologies, and performance metrics. Hydrology, 12(2), Article 39. https://doi.org/10.3390/hydrology12020039",

    "Ranjan, S., & Singh, V. P. (2022). HEC-HMS based rainfall-runoff model for Punpun river basin. Water Practice & "
    "Technology, 17(5), 986\u20131001. https://doi.org/10.2166/wpt.2022.033",

    "Sahu, M. K., Shwetha, H. R., & Dwarakish, G. S. (2023). State-of-the-art hydrological models and application of the "
    "HEC-HMS model: A review. Modeling Earth Systems and Environment, 9(3), 3029\u20133051. "
    "https://doi.org/10.1007/s40808-023-01704-7",

    "Shi, W., & Wang, N. (2020). An improved SCS-CN method incorporating slope, soil moisture, and storm duration factors "
    "for runoff prediction. Water, 12(5), Article 1335. https://doi.org/10.3390/w12051335",

    "Soulis, K. X. (2021). Soil Conservation Service Curve Number (SCS-CN) method: Current applications, remaining "
    "challenges, and future perspectives. Water, 13(2), Article 192. https://doi.org/10.3390/w13020192",

    "Souley Tangam, I., Yonaba, R., Niang, D., Adamou, M. M., Ke\u00efta, A., & Karambiri, H. (2024). Daily simulation of "
    "the rainfall\u2013runoff relationship in the Sirba River Basin in West Africa: Insights from the HEC-HMS model. "
    "Hydrology, 11(3), Article 34. https://doi.org/10.3390/hydrology11030034",

    "Tang, Y., Leon, A. S., & Kavvas, M. L. (2020). Impact of size and location of wetlands on watershed-scale flood "
    "control. Water Resources Management, 34(5), 1693\u20131707. https://doi.org/10.1007/s11269-020-02518-3",

    "Turkar, R. K., Shrivastava, R. N., Awasthi, M. K., & Lodhi, A. S. (2025). A comprehensive review of the HEC-HMS "
    "rainfall\u2013runoff simulation model and its hydrological applications. Journal of Agriculture and Ecology Research "
    "International, 26(5), 130\u2013144. https://doi.org/10.9734/jaeri/2025/v26i5711",

    "U.S. Army Corps of Engineers, Hydrologic Engineering Center. (n.d.). HEC-HMS technical reference manual. Retrieved "
    "August 17, 2026, from https://www.hec.usace.army.mil/confluence/hmsdocs/hmstrm",

    "Zhang, K., Shalehy, M. H., Ezaz, G. T., Chakraborty, A. K., Mohib, K. M., & Liu, L. (2022). An integrated flood risk "
    "assessment approach based on coupled hydrological-hydraulic modeling and bottom-up hazard vulnerability analysis. "
    "Environmental Modelling & Software, 148, Article 105279. https://doi.org/10.1016/j.envsoft.2021.105279",
]

for r_text in references:
    ref(r_text)

OUT = "thesis/Literature_Review_Gumai_Beel_Ichamoti_HEC-HMS.docx"
doc.save(OUT)

# --------------------------- verification: every body paragraph has citation
import re
uncited = []
for p in doc.paragraphs:
    t = p.text.strip()
    if not t or len(t.split()) < 25:          # skip headings/equation/short lines
        continue
    if t.startswith(("2.", "Q =")):
        continue
    if not re.search(r"\((?:[A-Z\u0100-\u017F][^()]*?, )?(?:n\.d\.|20\d\d)\)", t):
        uncited.append(t[:80])
words = sum(len(p.text.split()) for p in doc.paragraphs)
print(f"Saved {OUT}")
print(f"Approximate word count: {words}")
print(f"Paragraphs without citation: {len(uncited)}")
for u in uncited:
    print("  -", u)
