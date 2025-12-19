library(shiny)
library(shinythemes)
library(officer)
library(flextable)
library(magrittr)
library(pwr)
library(e1071)
library(tidyverse)
library(DiagrammeR)
library(rsvg)
library(htmltools)
library(shinyjs)
library(colourpicker)
library(V8)

counter_file <- "counter.rds"
if (!file.exists(counter_file)) saveRDS(0, counter_file)
counter <- readRDS(counter_file) + 1
saveRDS(counter, counter_file)


# Enhanced Professional CSS with animations
who_css <- "
/* Enhanced Professional Color Scheme */
:root {
  --who-blue: #0092D0;
  --who-light-blue: #6BC1E0;
  --who-dark-blue: #00689D;
  --who-green: #7CC242;
  --who-gray: #6D6E71;
  --who-light-gray: #F1F1F2;
  --professional-gold: #D4AF37;
  --professional-silver: #C0C0C0;
  --professional-bronze: #CD7F32;
}

/* Enhanced animations */
@keyframes float {
  0%, 100% { transform: translateY(0px); }
  50% { transform: translateY(-10px); }
}

@keyframes shimmer {
  0% { background-position: -1000px 0; }
  100% { background-position: 1000px 0; }
}

@keyframes pulse {
  0%, 100% { opacity: 1; }
  50% { opacity: 0.8; }
}

/* Enhanced professional header */
.navbar {
  background: linear-gradient(135deg, #00689D 0%, #0092D0 100%) !important;
  border-bottom: 3px solid var(--professional-gold) !important;
  box-shadow: 0 2px 10px rgba(0,0,0,0.1);
}

.navbar .navbar-nav > li > a {
  color: white !important;
  font-weight: 500;
  font-size: 15px;
  transition: all 0.3s ease;
  border-radius: 4px;
  margin: 2px 5px;
}

.navbar .navbar-nav > li > a:hover {
  background-color: rgba(255,255,255,0.2) !important;
  color: white !important;
  transform: translateY(-1px);
}

.navbar .navbar-brand {
  color: white !important;
  font-weight: bold;
  font-size: 20px;
  display: flex;
  align-items: center;
  gap: 10px;
}

/* Enhanced panels and cards */
.panel {
  border: none;
  border-radius: 8px;
  box-shadow: 0 2px 15px rgba(0,0,0,0.08);
  transition: all 0.3s ease;
  border: 1px solid #e0e0e0;
}

.panel:hover {
  box-shadow: 0 4px 25px rgba(0,0,0,0.12);
  transform: translateY(-2px);
}

.panel-heading {
  background: linear-gradient(135deg, var(--who-blue) 0%, var(--who-dark-blue) 100%) !important;
  color: white !important;
  border: none;
  border-radius: 8px 8px 0 0 !important;
  font-weight: 600;
  font-size: 16px;
  padding: 15px 20px;
}

.who-info-box {
  background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
  border-left: 4px solid var(--who-blue);
  padding: 20px;
  margin-bottom: 25px;
  border-radius: 8px;
  box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}

.who-formula-box {
  background-color: #F8F9FA;
  border: 1px solid #DEE2E6;
  padding: 20px;
  margin-bottom: 25px;
  border-radius: 8px;
  font-family: 'Courier New', monospace;
  border-left: 4px solid var(--professional-gold);
}

/* Enhanced buttons */
.btn-who {
  background: linear-gradient(135deg, var(--who-blue) 0%, var(--who-dark-blue) 100%);
  color: white;
  border: none;
  border-radius: 6px;
  font-weight: 600;
  padding: 10px 20px;
  transition: all 0.3s ease;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}

.btn-who:hover {
  background: linear-gradient(135deg, var(--who-dark-blue) 0%, #004A70 100%);
  color: white;
  transform: translateY(-2px);
  box-shadow: 0 4px 15px rgba(0,0,0,0.2);
}

.btn-who-secondary {
  background: linear-gradient(135deg, var(--who-green) 0%, #6BA83A 100%);
  color: white;
  border: none;
  border-radius: 6px;
  font-weight: 600;
  padding: 10px 20px;
  transition: all 0.3s ease;
}

.btn-who-secondary:hover {
  background: linear-gradient(135deg, #6BA83A 0%, #5A8C2E 100%);
  color: white;
  transform: translateY(-2px);
}

/* Enhanced footer */
.who-footer {
  background: linear-gradient(135deg, var(--who-gray) 0%, #5D5E60 100%);
  color: white;
  padding: 25px 0;
  text-align: center;
  margin-top: 40px;
  border-top: 3px solid var(--professional-gold);
}

.who-footer a {
  color: var(--who-light-blue);
  text-decoration: none;
  font-weight: 600;
  transition: all 0.3s ease;
}

.who-footer a:hover {
  color: white;
  text-decoration: underline;
}

/* Enhanced form controls */
.form-control {
  border-radius: 6px;
  border: 1px solid #CED4DA;
  padding: 10px 15px;
  transition: all 0.3s ease;
  font-size: 14px;
}

.form-control:focus {
  border-color: var(--who-light-blue);
  box-shadow: 0 0 0 0.3rem rgba(0, 146, 208, 0.15);
  transform: translateY(-1px);
}

/* Enhanced calculation sections */
.shiny-text-output, 
.verbatimTextOutput,
pre {
  background-color: white !important;
  color: black !important;
  padding: 25px !important;
  border-radius: 8px !important;
  border: 2px solid #E8F4FC !important;
  font-family: 'Arial', sans-serif !important;
  font-size: 16px !important;
  line-height: 1.6 !important;
  margin-bottom: 25px !important;
  white-space: pre-wrap !important;
  box-shadow: 0 2px 12px rgba(0,0,0,0.08) !important;
  border-left: 4px solid var(--who-blue) !important;
}

/* Enhanced table styling */
.table-responsive table {
  background-color: white !important;
  color: black !important;
  border: 2px solid #E8F4FC !important;
  font-size: 14px !important;
  margin: 25px 0 !important;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}

.table-responsive th {
  background: linear-gradient(135deg, var(--who-light-blue) 0%, var(--who-blue) 100%) !important;
  color: black !important;
  font-size: 14px !important;
  font-weight: 600 !important;
  padding: 15px !important;
  text-align: center !important;
  border: none !important;
}

.table-responsive td {
  background-color: white !important;
  color: black !important;
  border: 1px solid #e9ecef !important;
  font-size: 14px !important;
  padding: 12px 15px !important;
}

/* Enhanced headings */
h3, h4 {
  color: var(--who-dark-blue) !important;
  margin-top: 30px !important;
  margin-bottom: 20px !important;
  font-weight: 600;
}

h3 {
  font-size: 24px !important;
  border-bottom: 2px solid var(--professional-gold);
  padding-bottom: 10px;
}

h4 {
  font-size: 20px !important;
}


/* Home page animations */
.welcome-hero {
  background: linear-gradient(135deg, #0092D0 0%, #00689D 100%);
  color: white;
  padding: 80px 20px;
  text-align: center;
  border-radius: 15px;
  margin-bottom: 40px;
  position: relative;
  overflow: hidden;
}

.welcome-hero::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: linear-gradient(45deg, transparent 30%, rgba(255,255,255,0.1) 50%, transparent 70%);
  animation: shimmer 3s infinite linear;
}

.animated-features {
  display: flex;
  justify-content: center;
  flex-wrap: wrap;
  gap: 25px;
  margin-top: 40px;
}

.feature-box {
  background: rgba(255,255,255,0.15);
  backdrop-filter: blur(10px);
  padding: 25px;
  border-radius: 12px;
  width: 220px;
  border: 1px solid rgba(255,255,255,0.2);
  transition: all 0.3s ease;
  text-align: center;
}

.feature-box:hover {
  background: rgba(255,255,255,0.25);
  transform: translateY(-5px) scale(1.05);
  box-shadow: 0 10px 30px rgba(0,0,0,0.2);
}

.feature-box h4 {
  color: white !important;
  margin-bottom: 10px;
  font-weight: 600;
}

.feature-box p {
  color: rgba(255,255,255,0.9);
  font-size: 14px;
  line-height: 1.4;
}

/* Professional contact section */
.professional-contact {
  background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
  border-left: 4px solid var(--professional-gold);
  padding: 30px;
  margin: 40px 0;
  border-radius: 10px;
  box-shadow: 0 4px 15px rgba(0,0,0,0.08);
}

.professional-contact h4 {
  color: var(--who-dark-blue);
  margin-bottom: 20px;
  font-weight: 700;
  font-size: 22px;
}

.professional-contact p {
  font-size: 16px;
  line-height: 1.7;
  color: #495057;
  margin-bottom: 15px;
}

.professional-contact .contact-email {
  background: linear-gradient(135deg, var(--professional-gold) 0%, var(--professional-bronze) 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
  font-weight: 700;
  font-size: 18px;
  padding: 10px 0;
  display: inline-block;
}

/* Enhanced responsive design */
@media (max-width: 768px) {
  .sidebar-panel {
    margin-bottom: 20px;
  }
  
  .shiny-text-output, 
  .verbatimTextOutput,
  pre {
    padding: 20px !important;
    font-size: 14px !important;
    margin-bottom: 20px !important;
  }
  
  .table-responsive th,
  .table-responsive td {
    font-size: 12px !important;
    padding: 10px !important;
  }
  
  h3 {
    font-size: 20px !important;
  }
  
  h4 {
    font-size: 18px !important;
  }
  
  .welcome-header {
    font-size: 12px !important;
    padding: 10px 15px !important;
    text-align: center;
  }
  
  .animated-features {
    gap: 15px;
  }
  
  .feature-box {
    width: 100%;
    max-width: 280px;
  }
}

/* Additional professional enhancements */
.container-fluid {
  padding: 0 25px !important;
}

.main-panel {
  padding: 25px !important;
}

/* Highlight important numbers */
strong {
  color: var(--who-dark-blue) !important;
  font-weight: 700 !important;
}

em {
  color: #2C3E50 !important;
  font-style: italic !important;
}

/* Loading animations */
.loading-spinner {
  border: 4px solid #f3f3f3;
  border-top: 4px solid var(--who-blue);
  border-radius: 50%;
  width: 40px;
  height: 40px;
  animation: spin 2s linear infinite;
  margin: 20px auto;
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

/* Success states */
.success-message {
  background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
  border-left: 4px solid #28a745;
  color: #155724;
  padding: 15px;
  border-radius: 6px;
  margin: 10px 0;
}

/* Error states */
.error-message {
  background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
  border-left: 4px solid #dc3545;
  color: #721c24;
  padding: 15px;
  border-radius: 6px;
  margin: 10px 0;
}
"

ui <- navbarPage(
  title = div(icon("chart-bar"), "CalcuStats"),
  theme = shinytheme("flatly"),
  header = tags$head(
    tags$style(HTML(who_css))
  ),
  
  
  # Enhanced Home Tab with Professional Design
  tabPanel("Home",
           div(class = "container-fluid",
               # Animated Hero Section
               div(class = "row",
                   div(class = "col-md-12",
                       div(class = "welcome-hero",
                           h1("Welcome to CalcuStats", 
                              style = "font-size: 3.5em; font-weight: bold; margin-bottom: 20px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);"),
                           p("Your Comprehensive Statistical Companion for Research Excellence",
                             style = "font-size: 1.5em; margin-bottom: 40px; font-weight: 300;"),
                           
                           # Animated Feature Boxes
                           div(class = "animated-features",
                               div(class = "feature-box", style = "animation: float 3s ease-in-out infinite;",
                                   h4("Sample Size Calculators"),
                                   p("Precise calculations for various study designs with professional reporting")
                               ),
                               div(class = "feature-box", style = "animation: float 3s ease-in-out infinite 0.5s;",
                                   h4("Power Analysis"),
                                   p("Optimize your study's statistical power with robust calculations")
                               ),
                               div(class = "feature-box", style = "animation: float 3s ease-in-out infinite 1s;",
                                   h4("Descriptive Statistics"),
                                   p("Comprehensive data analysis and professional visualization")
                               ),
                               div(class = "feature-box", style = "animation: float 3s ease-in-out infinite 1.5s;",
                                   h4("Professional Reports"),
                                   p("Download detailed reports and visualizations for publications")
                               )
                           )
                       )
                   )
               ),
               
               # Main Content Area
               div(class = "row",
                   div(class = "col-md-8 col-md-offset-2",
                       div(class = "jumbotron", style = "background: linear-gradient(135deg, #E8F4FC 0%, #D4E8F7 100%); padding: 40px; border-radius: 15px;",
                           h2("Sample Size Calculation, Power Analysis, and Descriptive Analysis Made Simple", style = "color: #00689D; text-align: center;"),
                           p("CalcuStats provides researchers, healthcare professionals, and students with reliable, validated methods for sample size determination, power analysis, and descriptive statistics (Developed by Mudasir Mohammed Ibrahim, BSc, RN).", 
                             style = "font-size: 16px; line-height: 1.6; text-align: center; margin-bottom: 30px;"),
                           hr(style = "border-color: #0092D0;"),
                           p("This application combines statistical rigor with user-friendly design to support evidence-based research and decision-making.", 
                             style = "font-size: 15px; line-height: 1.6;")
                       ),
                       
                       # Feature Cards
                       div(class = "row",
                           div(class = "col-md-4",
                               div(class = "panel panel-default",
                                   div(class = "panel-heading", 
                                       h3("Sample Size Calculators", class = "panel-title")
                                   ),
                                   div(class = "panel-body",
                                       p("Determine appropriate sample sizes using validated methods:"),
                                       tags$ul(
                                         tags$li("Taro Yamane formula"),
                                         tags$li("Cochran's formula"),
                                         tags$li("Single Population Proportion"),
                                         tags$li("Case-Control Studies"),
                                         tags$li("Cohort Studies"),
                                         tags$li("Various other statistical formulas")
                                       )
                                   )
                               )
                           ),
                           div(class = "col-md-4",
                               div(class = "panel panel-default",
                                   div(class = "panel-heading", 
                                       h3("Power Analysis", class = "panel-title")
                                   ),
                                   div(class = "panel-body",
                                       p("Calculate statistical power for various study designs:"),
                                       tags$ul(
                                         tags$li("t-tests (independent, paired, one-sample)"),
                                         tags$li("ANOVA (one-way, two-way)"),
                                         tags$li("Correlation studies"),
                                         tags$li("Regression analysis"),
                                         tags$li("Chi-square tests"),
                                         tags$li("Proportion comparisons")
                                       )
                                   )
                               )
                           ),
                           div(class = "col-md-4",
                               div(class = "panel panel-default",
                                   div(class = "panel-heading", 
                                       h3("Descriptive Statistics", class = "panel-title")
                                   ),
                                   div(class = "panel-body",
                                       p("Generate comprehensive descriptive statistics:"),
                                       tags$ul(
                                         tags$li("Central tendency measures"),
                                         tags$li("Dispersion metrics"),
                                         tags$li("Distribution characteristics"),
                                         tags$li("Data visualization"),
                                         tags$li("Professional reporting")
                                       )
                                   )
                               )
                           )
                       ),
                       
                       # Professional Contact Section
                       div(class = "professional-contact",
                           h4("Contribute to CalcuStats Continuous Development"),
                           p("We are continuously working to enhance CalcuStats with new features and formulas. Your input is valuable in making this tool more comprehensive and useful for the research community."),
                           p("If you have suggestions for additional formulas or statistical methods that could benefit researchers, or if you encounter any issues, please don't hesitate to share them with us."),
                           p("We welcome:"),
                           tags$ul(
                             tags$li("New formula suggestions for sample size calculation"),
                             tags$li("Any feedback to improve user experience")
                           ),
                           p("Contact us at:"),
                           div(class = "contact-email", "mudassiribrahim30@gmail.com"),
                           p("Your contributions help make CalcuStats a better tool for everyone in the research community.")
                       )
                   )
               )
           )
  ),
  
  # Proportional Allocation Tab (unchanged but with enhanced styling)
  tabPanel("Proportional Allocation",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "proportional_sidebar",
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Proportional Allocation Parameters"),
                 div(class = "panel-body",
                     div(
                       style = "margin-bottom: 15px;",
                       actionButton("reset_proportional", "Reset All Values", class = "btn-who",
                                    icon = icon("refresh"))
                     ),
                     numericInput("custom_sample", "Your Sample Size (n)", value = 100, min = 1, step = 1),
                     numericInput("non_response_custom", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
                     actionButton("addStratum_custom", "Add Stratum", class = "btn-who"),
                     actionButton("removeStratum_custom", "Remove Last Stratum", class = "btn-who-secondary"),
                     uiOutput("stratumInputs_custom"),
                     div(
                       class = "who-info-box",
                       style = "margin-top: 10px; font-size: 14px;",
                       p("For a single population, enter the total size below."),
                       p("For multiple populations, click 'Add Stratum' to create additional groups.")
                     )
                 )
               ),
               
               conditionalPanel(
                 condition = "output.stratumCount_custom > 1",
                 div(
                   class = "panel panel-default",
                   div(class = "panel-heading", "Flow Chart Options"),
                   div(class = "panel-body",
                       checkboxInput("showFlowchart_custom", "Generate Flow Chart Diagram", FALSE),
                       conditionalPanel(
                         condition = "input.showFlowchart_custom",
                         div(
                           style = "max-height: 400px; overflow-y: auto;",
                           selectInput("flowchartLayout_custom", "Layout Style:",
                                       choices = c("dot", "neato", "twopi", "circo", "fdp")),
                           colourpicker::colourInput("nodeColor_custom", "Node Color", value = "#6BAED6"),
                           colourpicker::colourInput("edgeColor_custom", "Edge Color", value = "#636363"),
                           sliderInput("nodeFontSize_custom", "Default Node Font Size", min = 12, max = 24, value = 16),
                           sliderInput("edgeFontSize_custom", "Default Edge Font Size", min = 10, max = 20, value = 14),
                           sliderInput("nodeWidth_custom", "Node Width", min = 0.5, max = 3, value = 1, step = 0.1),
                           sliderInput("nodeHeight_custom", "Node Height", min = 0.5, max = 3, value = 0.8, step = 0.1),
                           selectInput("nodeShape_custom", "Node Shape", 
                                       choices = c("rectangle", "ellipse", "circle", "diamond", "triangle", "hexagon"),
                                       selected = "rectangle"),
                           sliderInput("arrowSize_custom", "Arrow Size", min = 0.1, max = 2, value = 1, step = 0.1),
                           
                           # New options for including stratum/pop sizes
                           h4("Diagram Content Options"),
                           checkboxInput("includePopSize_custom", "Include Population Size", value = TRUE),
                           checkboxInput("includeStratumSize_custom", "Include Stratum Size", value = TRUE),
                           
                           actionButton("updateFlowchart_custom", "Update Diagram", class = "btn-who"),
                           br(), br(),
                           h4("Diagram Editor"),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;",
                             h5("Text Editing"),
                             textInput("nodeText_custom", "Current Text:", ""),
                             textInput("newText_custom", "New Text (use -> for replacement):", ""),
                             actionButton("editNodeText_custom", "Update Text", class = "btn-who"),
                             actionButton("deleteText_custom", "Delete Text", class = "btn-who-secondary"),
                             actionButton("resetDiagramText_custom", "Reset All Text", class = "btn-warning")
                           ),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px;",
                             h5("Font Size Controls"),
                             sliderInput("selectedNodeFontSize_custom", "Selected Node Font Size", 
                                         min = 12, max = 24, value = 16),
                             sliderInput("selectedEdgeFontSize_custom", "Selected Edge Font Size", 
                                         min = 10, max = 20, value = 14),
                             actionButton("applyFontSizes_custom", "Apply Font Sizes", class = "btn-who")
                           ),
                           br(), br(),
                           h4("Download Diagram"),
                           div(style = "display: inline-block;", 
                               downloadButton("downloadFlowchartPNG_custom", "PNG (High Quality)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartSVG_custom", "SVG (Vector)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartPDF_custom", "PDF (Vector)", class = "btn-who-secondary"))
                         )
                       )
                   )
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Export Results"),
                 div(class = "panel-body",
                     downloadButton("downloadCustomWord", "Download as Word", class = "btn-who"),
                     downloadButton("downloadCustomSteps", "Download Calculation Steps", class = "btn-who-secondary")
                 )
               )
             ),
             mainPanel(
               width = 8,
               h3("Proportional Allocation for Specified Sample Size"),
               div(
                 class = "who-info-box",
                 HTML("
            <strong>Proportional Allocation Formula:</strong><br>
            <em>n<sub>i</sub> = (N<sub>i</sub> / N) &times; n</em><br>
            <ul>
              <li><strong>n<sub>i</sub></strong> = sample size for stratum i</li>
              <li><strong>N<sub>i</sub></strong> = population size for stratum i</li>
              <li><strong>N</strong> = total population size</li>
              <li><strong>n</strong> = total sample size</li>
            </ul>
            This method distributes your specified sample size proportionally across strata based on their population sizes.
          ")
               ),
               fluidRow(
                 column(6, div(style = "font-size: 14px;", verbatimTextOutput("totalPopCustom"))),
                 column(6, div(style = "font-size: 14px;", verbatimTextOutput("sampleSizeCustom")))
               ),
               div(style = "font-size: 14px;", verbatimTextOutput("adjustedSampleSizeCustom")),
               h4("Proportional Allocation"),
               div(class = "table-responsive",
                   tableOutput("allocationTableCustom")
               ),
               div(style = "font-size: 14px;", textOutput("interpretationTextCustom")),
               conditionalPanel(
                 condition = "input.showFlowchart_custom && output.stratumCount_custom > 1",
                 h4("Stratification Flow Chart"),
                 grVizOutput("flowchart_custom", width = "100%", height = "500px"),
                 div(id = "diagramEditor_custom",
                     style = "margin-top: 20px; border: 1px solid #ddd; padding: 10px; font-size: 14px;",
                     h4("Interactive Diagram Editor"),
                     p("Click on nodes or edges in the diagram to edit them."),
                     uiOutput("nodeEdgeEditorUI_custom")
                 )
               )
             )
           )
  ),
  
  # Taro Yamane Tab (unchanged)
  tabPanel("Taro Yamane",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "yamane_sidebar",
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Taro Yamane Parameters"),
                 div(class = "panel-body",
                     div(
                       style = "margin-bottom: 15px;",
                       actionButton("reset_yamane", "Reset All Values", class = "btn-who",
                                    icon = icon("refresh"))
                     ),
                     numericInput("e", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
                     numericInput("non_response", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
                     actionButton("addStratum", "Add Stratum", class = "btn-who"),
                     actionButton("removeStratum", "Remove Last Stratum", class = "btn-who-secondary"),
                     uiOutput("stratumInputs"),
                     div(
                       class = "who-info-box",
                       style = "margin-top: 10px; font-size: 14px;",
                       p("For a single population, enter the total size below."),
                       p("For multiple populations, click 'Add Stratum' to create additional groups.")
                     )
                 )
               ),
               
               conditionalPanel(
                 condition = "output.stratumCount > 1",
                 div(
                   class = "panel panel-default",
                   div(class = "panel-heading", "Flow Chart Options"),
                   div(class = "panel-body",
                       checkboxInput("showFlowchart", "Generate Flow Chart Diagram", FALSE),
                       conditionalPanel(
                         condition = "input.showFlowchart",
                         div(
                           style = "max-height: 400px; overflow-y: auto;",
                           selectInput("flowchartLayout", "Layout Style:",
                                       choices = c("dot", "neato", "twopi", "circo", "fdp")),
                           colourpicker::colourInput("nodeColor", "Node Color", value = "#6BAED6"),
                           colourpicker::colourInput("edgeColor", "Edge Color", value = "#636363"),
                           sliderInput("nodeFontSize", "Default Node Font Size", min = 12, max = 24, value = 16),
                           sliderInput("edgeFontSize", "Default Edge Font Size", min = 10, max = 20, value = 14),
                           sliderInput("nodeWidth", "Node Width", min = 0.5, max = 3, value = 1, step = 0.1),
                           sliderInput("nodeHeight", "Node Height", min = 0.5, max = 3, value = 0.8, step = 0.1),
                           selectInput("nodeShape", "Node Shape", 
                                       choices = c("rectangle", "ellipse", "circle", "diamond", "triangle", "hexagon"),
                                       selected = "rectangle"),
                           sliderInput("arrowSize", "Arrow Size", min = 0.1, max = 2, value = 1, step = 0.1),
                           
                           # New options for including stratum/pop sizes
                           h4("Diagram Content Options"),
                           checkboxInput("includePopSize", "Include Population Size", value = TRUE),
                           checkboxInput("includeStratumSize", "Include Stratum Size", value = TRUE),
                           
                           actionButton("updateFlowchart", "Update Diagram", class = "btn-who"),
                           br(), br(),
                           h4("Diagram Editor"),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;",
                             h5("Text Editing"),
                             textInput("nodeText", "Current Text:", ""),
                             textInput("newText", "New Text (use -> for replacement):", ""),
                             actionButton("editNodeText", "Update Text", class = "btn-who"),
                             actionButton("deleteText", "Delete Text", class = "btn-who-secondary"),
                             actionButton("resetDiagramText", "Reset All Text", class = "btn-warning")
                           ),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px;",
                             h5("Font Size Controls"),
                             sliderInput("selectedNodeFontSize", "Selected Node Font Size", 
                                         min = 12, max = 24, value = 16),
                             sliderInput("selectedEdgeFontSize", "Selected Edge Font Size", 
                                         min = 10, max = 20, value = 14),
                             actionButton("applyFontSizes", "Apply Font Sizes", class = "btn-who")
                           ),
                           br(), br(),
                           h4("Download Diagram"),
                           div(style = "display: inline-block;", 
                               downloadButton("downloadFlowchartPNG", "PNG (High Quality)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartSVG", "SVG (Vector)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartPDF", "PDF (Vector)", class = "btn-who-secondary"))
                         )
                       )
                   )
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Export Results"),
                 div(class = "panel-body",
                     downloadButton("downloadWord", "Download as Word", class = "btn-who"),
                     downloadButton("downloadSteps", "Download Calculation Steps", class = "btn-who-secondary")
                 )
               )
             ),
             mainPanel(
               width = 8,
               h3("Taro Yamane Method with Proportional Allocation"),
               div(
                 class = "who-info-box",
                 HTML("
            <strong>Taro Yamane Formula:</strong><br>
            <em>n = N / (1 + N &times; e²)</em><br>
            <ul>
              <li><strong>n</strong> = required sample size</li>
              <li><strong>N</strong> = population size</li>
              <li><strong>e</strong> = margin of error (in decimal)</li>
            </ul>
            This formula assumes a 95% confidence level and simplifies sample size estimation for large populations.
          ")
               ),
               fluidRow(
                 column(6, div(style = "font-size: 14px;", verbatimTextOutput("totalPop"))),
                 column(6, div(style = "font-size: 14px;", verbatimTextOutput("sampleSize")))
               ),
               div(style = "font-size: 14px;", verbatimTextOutput("formulaExplanation")),
               div(style = "font-size: 14px;", verbatimTextOutput("adjustedSampleSize")),
               h4("Proportional Allocation"),
               div(class = "table-responsive",
                   tableOutput("allocationTable")
               ),
               div(style = "font-size: 14px;", textOutput("interpretationText")),
               conditionalPanel(
                 condition = "input.showFlowchart && output.stratumCount > 1",
                 h4("Stratification Flow Chart"),
                 grVizOutput("flowchart", width = "100%", height = "500px"),
                 div(id = "diagramEditor",
                     style = "margin-top: 20px; border: 1px solid #ddd; padding: 10px; font-size: 14px;",
                     h4("Interactive Diagram Editor"),
                     p("Click on nodes or edges in the diagram to edit them."),
                     uiOutput("nodeEdgeEditorUI")
                 )
               )
             )
           )
  ),
  
  # UPDATED: Cochran Formula Tab with percentage input support
  tabPanel("Cochran Formula",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "cochran_sidebar",
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Cochran Formula Parameters"),
                 div(class = "panel-body",
                     div(
                       style = "margin-bottom: 15px;",
                       actionButton("reset_cochran", "Reset All Values", class = "btn-who",
                                    icon = icon("refresh"))
                     ),
                     
                     # UPDATED: Proportion input with percentage support
                     div(
                       style = "margin-bottom: 15px;",
                       textInput("p_input", "Estimated Proportion (p) or Percentage (%)", 
                                 value = "0.5",
                                 placeholder = "Enter as proportion (0.5) or percentage (50%)"),
                       div(
                         class = "who-info-box",
                         style = "margin-top: 5px; font-size: 12px;",
                         HTML("<strong>Flexible Input Format:</strong><br>
                              • Enter as proportion: <em>0.5</em> (for 50%)<br>
                              • Enter as percentage: <em>50%</em> or <em>50</em><br>
                              • Both will give same results<br>
                              <em>Examples: 0.3, 30%, 5%, 0.05</em>")
                       )
                     ),
                     
                     numericInput("z", "Z-score (Z)", value = 1.96),
                     numericInput("e_c", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
                     numericInput("N_c", "Population Size (optional)", value = NULL, min = 1),
                     numericInput("non_response_c", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
                     actionButton("addStratum_c", "Add Stratum", class = "btn-who"),
                     actionButton("removeStratum_c", "Remove Last Stratum", class = "btn-who-secondary"),
                     uiOutput("stratumInputs_c")
                 )
               ),
               
               conditionalPanel(
                 condition = "output.stratumCount_c > 1",
                 div(
                   class = "panel panel-default",
                   div(class = "panel-heading", "Flow Chart Options"),
                   div(class = "panel-body",
                       checkboxInput("showFlowchart_c", "Generate Flow Chart Diagram", FALSE),
                       conditionalPanel(
                         condition = "input.showFlowchart_c",
                         div(
                           style = "max-height: 400px; overflow-y: auto;",
                           selectInput("flowchartLayout_c", "Layout Style:",
                                       choices = c("dot", "neato", "twopi", "circo", "fdp")),
                           colourpicker::colourInput("nodeColor_c", "Node Color", value = "#6BAED6"),
                           colourpicker::colourInput("edgeColor_c", "Edge Color", value = "#636363"),
                           sliderInput("nodeFontSize_c", "Default Node Font Size", min = 12, max = 24, value = 16),
                           sliderInput("edgeFontSize_c", "Default Edge Font Size", min = 10, max = 20, value = 14),
                           sliderInput("nodeWidth_c", "Node Width", min = 0.5, max = 3, value = 1, step = 0.1),
                           sliderInput("nodeHeight_c", "Node Height", min = 0.5, max = 3, value = 0.8, step = 0.1),
                           selectInput("nodeShape_c", "Node Shape", 
                                       choices = c("rectangle", "ellipse", "circle", "diamond", "triangle", "hexagon"),
                                       selected = "rectangle"),
                           sliderInput("arrowSize_c", "Arrow Size", min = 0.1, max = 2, value = 1, step = 0.1),
                           
                           # New options for including stratum/pop sizes
                           h4("Diagram Content Options"),
                           checkboxInput("includePopSize_c", "Include Population Size", value = TRUE),
                           checkboxInput("includeStratumSize_c", "Include Stratum Size", value = TRUE),
                           
                           actionButton("updateFlowchart_c", "Update Diagram", class = "btn-who"),
                           br(), br(),
                           h4("Diagram Editor"),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;",
                             h5("Text Editing"),
                             textInput("nodeText_c", "Current Text:", ""),
                             textInput("newText_c", "New Text (use -> for replacement):", ""),
                             actionButton("editNodeText_c", "Update Text", class = "btn-who"),
                             actionButton("deleteText_c", "Delete Text", class = "btn-who-secondary"),
                             actionButton("resetDiagramText_c", "Reset All Text", class = "btn-warning")
                           ),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px;",
                             h5("Font Size Controls"),
                             sliderInput("selectedNodeFontSize_c", "Selected Node Font Size", 
                                         min = 12, max = 24, value = 16),
                             sliderInput("selectedEdgeFontSize_c", "Selected Edge Font Size", 
                                         min = 10, max = 20, value = 14),
                             actionButton("applyFontSizes_c", "Apply Font Sizes", class = "btn-who")
                           ),
                           br(), br(),
                           h4("Download Diagram"),
                           div(style = "display: inline-block;", 
                               downloadButton("downloadFlowchartPNG_c", "PNG (High Quality)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartSVG_c", "SVG (Vector)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartPDF_c", "PDF (Vector)", class = "btn-who-secondary"))
                         )
                       )
                   )
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Export Results"),
                 div(class = "panel-body",
                     downloadButton("downloadCochranWord", "Download as Word", class = "btn-who"),
                     downloadButton("downloadCochranSteps", "Download Calculation Steps", class = "btn-who-secondary")
                 )
               )
             ),
             mainPanel(
               width = 8,
               h3("Cochran Sample Size Calculator with Proportional Allocation"),
               div(
                 class = "who-info-box",
                 HTML("
            <strong>Cochran's Formula:</strong><br>
            <em>n₀ = (Z² × p × (1-p)) / e²</em><br>
            <ul>
              <li><strong>n₀</strong> = required sample size</li>
              <li><strong>Z</strong> = Z-score for desired confidence level</li>
              <li><strong>p</strong> = estimated proportion of the attribute present in the population</li>
              <li><strong>e</strong> = margin of error (in decimal)</li>
            </ul>
            For finite populations, apply the finite population correction: n = n₀ / (1 + (n₀ - 1)/N)
          ")
               ),
               div(style = "font-size: 14px;", verbatimTextOutput("cochranFormulaExplanation")),
               div(style = "font-size: 14px;", verbatimTextOutput("cochranSample")),
               div(style = "font-size: 14px;", verbatimTextOutput("cochranAdjustedSample")),
               h4("Proportional Allocation"),
               div(class = "table-responsive",
                   tableOutput("cochranAllocationTable")
               ),
               div(style = "font-size: 14px;", textOutput("cochranInterpretation")),
               conditionalPanel(
                 condition = "input.showFlowchart_c && output.stratumCount_c > 1",
                 h4("Stratification Flow Chart"),
                 grVizOutput("flowchart_c", width = "100%", height = "500px"),
                 div(id = "diagramEditor_c",
                     style = "margin-top: 20px; border: 1px solid #ddd; padding: 10px; font-size: 14px;",
                     h4("Interactive Diagram Editor"),
                     p("Click on nodes or edges in the diagram to edit them."),
                     uiOutput("nodeEdgeEditorUI_c")
                 )
               )
             )
           )
  ),
  
  # UPDATED: Other Formulas Tab with Single Population Proportion
  tabPanel("Other Formulas",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "other_formulas_sidebar",
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Study Design Selection"),
                 div(class = "panel-body",
                     div(
                       style = "margin-bottom: 15px;",
                       actionButton("reset_other", "Reset All Values", class = "btn-who",
                                    icon = icon("refresh"))
                     ),
                     selectInput("formula_type", "Select Study Design:",
                                 choices = c("Single Population Proportion (EPI INFO)" = "single_proportion",
                                             "Mean Estimation (Known Variance)" = "mean_known_var",
                                             "Mean Estimation (Unknown Variance)" = "mean_unknown_var",
                                             "Proportion Difference" = "proportion_diff",
                                             "Correlation Coefficient" = "correlation",
                                             "Regression Coefficient" = "regression",
                                             "Odds Ratio" = "odds_ratio",
                                             "Relative Risk" = "relative_risk",
                                             "Prevalence Study" = "prevalence",
                                             "Case-Control Study" = "case_control",
                                             "Cohort Study" = "cohort")),
                     uiOutput("formula_params"),
                     numericInput("non_response_other", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1)
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Stratification"),
                 div(class = "panel-body",
                     actionButton("addStratum_other", "Add Stratum", class = "btn-who"),
                     actionButton("removeStratum_other", "Remove Last Stratum", class = "btn-who-secondary"),
                     uiOutput("stratumInputs_other")
                 )
               ),
               
               conditionalPanel(
                 condition = "output.stratumCount_other > 1",
                 div(
                   class = "panel panel-default",
                   div(class = "panel-heading", "Flow Chart Options"),
                   div(class = "panel-body",
                       checkboxInput("showFlowchart_other", "Generate Flow Chart Diagram", FALSE),
                       conditionalPanel(
                         condition = "input.showFlowchart_other",
                         div(
                           style = "max-height: 400px; overflow-y: auto;",
                           selectInput("flowchartLayout_other", "Layout Style:",
                                       choices = c("dot", "neato", "twopi", "circo", "fdp")),
                           colourpicker::colourInput("nodeColor_other", "Node Color", value = "#6BAED6"),
                           colourpicker::colourInput("edgeColor_other", "Edge Color", value = "#636363"),
                           sliderInput("nodeFontSize_other", "Default Node Font Size", min = 12, max = 24, value = 16),
                           sliderInput("edgeFontSize_other", "Default Edge Font Size", min = 10, max = 20, value = 14),
                           sliderInput("nodeWidth_other", "Node Width", min = 0.5, max = 3, value = 1, step = 0.1),
                           sliderInput("nodeHeight_other", "Node Height", min = 0.5, max = 3, value = 0.8, step = 0.1),
                           selectInput("nodeShape_other", "Node Shape", 
                                       choices = c("rectangle", "ellipse", "circle", "diamond", "triangle", "hexagon"),
                                       selected = "rectangle"),
                           sliderInput("arrowSize_other", "Arrow Size", min = 0.1, max = 2, value = 1, step = 0.1),
                           
                           h4("Diagram Content Options"),
                           checkboxInput("includePopSize_other", "Include Population Size", value = TRUE),
                           checkboxInput("includeStratumSize_other", "Include Stratum Size", value = TRUE),
                           
                           actionButton("updateFlowchart_other", "Update Diagram", class = "btn-who"),
                           br(), br(),
                           h4("Diagram Editor"),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;",
                             h5("Text Editing"),
                             textInput("nodeText_other", "Current Text:", ""),
                             textInput("newText_other", "New Text (use -> for replacement):", ""),
                             actionButton("editNodeText_other", "Update Text", class = "btn-who"),
                             actionButton("deleteText_other", "Delete Text", class = "btn-who-secondary"),
                             actionButton("resetDiagramText_other", "Reset All Text", class = "btn-warning")
                           ),
                           div(
                             style = "border: 1px solid #ddd; padding: 10px;",
                             h5("Font Size Controls"),
                             sliderInput("selectedNodeFontSize_other", "Selected Node Font Size", 
                                         min = 12, max = 24, value = 16),
                             sliderInput("selectedEdgeFontSize_other", "Selected Edge Font Size", 
                                         min = 10, max = 20, value = 14),
                             actionButton("applyFontSizes_other", "Apply Font Sizes", class = "btn-who")
                           ),
                           br(), br(),
                           h4("Download Diagram"),
                           div(style = "display: inline-block;", 
                               downloadButton("downloadFlowchartPNG_other", "PNG (High Quality)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartSVG_other", "SVG (Vector)", class = "btn-who-secondary")),
                           div(style = "display: inline-block; margin-left: 5px;", 
                               downloadButton("downloadFlowchartPDF_other", "PDF (Vector)", class = "btn-who-secondary"))
                         )
                       )
                   )
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Export Results"),
                 div(class = "panel-body",
                     downloadButton("downloadOtherWord", "Download as Word", class = "btn-who"),
                     downloadButton("downloadOtherSteps", "Download Calculation Steps", class = "btn-who-secondary")
                 )
               )
             ),
             mainPanel(
               width = 8,
               h3("Sample Size Calculation for Various Study Designs"),
               uiOutput("formula_description"),
               div(style = "font-size: 14px;", verbatimTextOutput("formulaExplanation_other")),
               div(style = "font-size: 14px;", verbatimTextOutput("sampleSize_other")),
               div(style = "font-size: 14px;", verbatimTextOutput("adjustedSampleSize_other")),
               h4("Proportional Allocation"),
               div(class = "table-responsive",
                   tableOutput("allocationTable_other")
               ),
               div(style = "font-size: 14px;", textOutput("interpretationText_other")),
               conditionalPanel(
                 condition = "input.showFlowchart_other && output.stratumCount_other > 1",
                 h4("Stratification Flow Chart"),
                 grVizOutput("flowchart_other", width = "100%", height = "500px"),
                 div(id = "diagramEditor_other",
                     style = "margin-top: 20px; border: 1px solid #ddd; padding = 10px; font-size: 14px;",
                     h4("Interactive Diagram Editor"),
                     p("Click on nodes or edges in the diagram to edit them."),
                     uiOutput("nodeEdgeEditorUI_other")
                 )
               )
             )
           )
  ),
  
  # UPDATED: Power Analysis Tab with Enhanced Simple Linear Regression Support
  tabPanel("Power Analysis",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "power_sidebar",
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Power Analysis Parameters"),
                 div(class = "panel-body",
                     div(
                       style = "margin-bottom: 15px;",
                       actionButton("reset_power", "Reset All Values", class = "btn-who",
                                    icon = icon("refresh"))
                     ),
                     selectInput("testType", "Statistical Test:",
                                 choices = c("Independent t-test", "Paired t-test", "One-sample t-test", "One-Way ANOVA",
                                             "Two-Way ANOVA", "Proportion", "Correlation", "Chi-squared",
                                             "Simple Linear Regression", "Multiple Linear Regression")),
                     
                     # Effect size input with dynamic labels
                     conditionalPanel(
                       condition = "input.testType == 'Simple Linear Regression' || input.testType == 'Multiple Linear Regression'",
                       numericInput("effectSize", "Effect Size (f²)", value = 0.15, min = 0.01, step = 0.01)
                     ),
                     conditionalPanel(
                       condition = "input.testType != 'Simple Linear Regression' && input.testType != 'Multiple Linear Regression'",
                       numericInput("effectSize", "Effect Size", value = 0.3, min = 0.01, step = 0.01)
                     ),
                     
                     numericInput("alpha", "Significance Level (alpha)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                     numericInput("power", "Desired Power", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                     
                     # Test-specific parameters
                     conditionalPanel(
                       condition = "input.testType == 'Multiple Linear Regression'",
                       numericInput("predictors", "Number of Predictors (u)", value = 2, min = 1, step = 1)
                     ),
                     conditionalPanel(
                       condition = "input.testType == 'Simple Linear Regression'",
                       div(class = "who-info-box",
                           style = "margin-top: 10px; font-size: 12px;",
                           HTML("<strong>Effect Size Guidelines (f²):</strong><br>
                              Small: 0.02<br>
                              Medium: 0.15<br>
                              Large: 0.35")
                       )
                     ),
                     conditionalPanel(
                       condition = "input.testType == 'One-Way ANOVA' || input.testType == 'Two-Way ANOVA'",
                       numericInput("groups", "Number of Groups", value = 3, min = 2)
                     ),
                     conditionalPanel(
                       condition = "input.testType == 'Chi-squared'",
                       numericInput("df", "Degrees of Freedom", value = 1, min = 1),
                       numericInput("nrow", "Number of Rows", value = 2, min = 2),
                       numericInput("ncol", "Number of Columns", value = 2, min = 2)
                     ),
                     actionButton("runPower", "Run Power Analysis", class = "btn-who")
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Effect Size Help"),
                 div(class = "panel-body",
                     conditionalPanel(
                       condition = "input.testType == 'Simple Linear Regression'",
                       HTML("
                       <strong>Simple Linear Regression Effect Size (f²):</strong><br>
                       f² = R² / (1 - R²)<br><br>
                       <strong>Where:</strong><br>
                       R² = proportion of variance explained<br><br>
                       <strong>Common values:</strong><br>
                       • Small: 0.02 (R² ≈ 2%)<br>
                       • Medium: 0.15 (R² ≈ 13%)<br>
                       • Large: 0.35 (R² ≈ 26%)
                     ")
                     )
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Export Results"),
                 div(class = "panel-body",
                     downloadButton("downloadPowerSteps", "Download Calculation Steps", class = "btn-who-secondary")
                 )
               )
             ),
             mainPanel(
               width = 8,
               h3("Power Analysis for Inferential Tests"),
               div(
                 class = "who-info-box",
                 HTML("
            <strong>Power Analysis:</strong><br>
            <p>Statistical power is the probability that a test will correctly reject a false null hypothesis (avoid a Type II error).</p>
            <p>Power is influenced by:</p>
            <ul>
              <li><strong>Effect size:</strong> The magnitude of the difference or relationship you want to detect</li>
              <li><strong>Sample size:</strong> The number of observations in your study</li>
              <li><strong>Significance level (α):</strong> The probability of a Type I error (false positive)</li>
            </ul>
            <p><strong>Effect Size Guidelines:</strong></p>
            <ul>
              <li><strong>t-tests:</strong> Cohen's d (small: 0.2, medium: 0.5, large: 0.8)</li>
              <li><strong>ANOVA:</strong> Cohen's f (small: 0.1, medium: 0.25, large: 0.4)</li>
              <li><strong>Correlation:</strong> r (small: 0.1, medium: 0.3, large: 0.5)</li>
              <li><strong>Chi-squared:</strong> w (small: 0.1, medium: 0.3, large: 0.5)</li>
              <li><strong>Simple Linear Regression:</strong> f² (small: 0.02, medium: 0.15, large: 0.35)</li>
              <li><strong>Multiple Linear Regression:</strong> f² (small: 0.02, medium: 0.15, large: 0.35)</li>
            </ul>
            <p>A power of 0.8 (80%) is generally considered acceptable in most research contexts.</p>
          ")
               ),
               div(style = "font-size: 14px;", verbatimTextOutput("powerResult")),
               conditionalPanel(
                 condition = "output.powerError",
                 div(class = "error-message",
                     textOutput("powerError")
                 )
               ),
               conditionalPanel(
                 condition = "input.testType == 'Simple Linear Regression'",
                 div(class = "who-info-box",
                     style = "margin-top: 20px;",
                     HTML("
                     <strong>Simple Linear Regression Power Analysis:</strong><br>
                     <p>This calculates the sample size needed to detect a significant relationship between a single predictor variable and an outcome variable.</p>
                     <p><strong>Formula:</strong> n = (Zα + Zβ)² / f² + 2</p>
                     <p><strong>Where:</strong></p>
                     <ul>
                       <li>Zα = Z-score for alpha (Type I error rate)</li>
                       <li>Zβ = Z-score for beta (Type II error rate)</li>
                       <li>f² = R² / (1 - R²)</li>
                       <li>n = total sample size</li>
                     </ul>
                     <p>The calculation accounts for the 2 parameters being estimated (intercept and slope).</p>
                   ")
                 )
               )
             )
           )
  ),
  
  # Descriptive Statistics Tab (unchanged)
  tabPanel("Descriptive Statistics",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "desc_sidebar",
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Data Input"),
                 div(class = "panel-body",
                     div(
                       style = "margin-bottom: 15px;",
                       actionButton("reset_desc", "Reset All Values", class = "btn-who",
                                    icon = icon("refresh"))
                     ),
                     tags$textarea(id = "dataInput", rows = 10, cols = 30,
                                   placeholder = "Paste your data (with header) from Excel or statistical software...",
                                   style = "font-size: 14px; width: 100%;"),
                     selectInput("dataType", "Data Type:",
                                 choices = c("Auto Detect" = "auto",
                                             "Numerical/Continuous" = "numerical",
                                             "Categorical (Nominal)" = "nominal",
                                             "Ordinal" = "ordinal")),
                     actionButton("runDesc", "Analyze Data", class = "btn-who")
                 )
               ),
               
               div(
                 class = "panel panel-default",
                 div(class = "panel-heading", "Export Results"),
                 div(class = "panel-body",
                     downloadButton("downloadDescSteps", "Download Full Report (Word)", class = "btn-who-secondary")
                 )
               )
             ),
             mainPanel(
               width = 8,
               h3("Descriptive Statistics"),
               div(
                 class = "who-info-box",
                 HTML("
          <strong>Descriptive Statistics:</strong><br>
          <p>Descriptive statistics summarize and describe the main features of a dataset. The appropriate statistics depend on the measurement scale:</p>
          <ul>
            <li><strong>Numerical/Continuous Data:</strong> Mean, median, standard deviation, range, etc.</li>
            <li><strong>Categorical (Nominal) Data:</strong> Frequency counts, percentages, mode</li>
            <li><strong>Ordinal Data:</strong> Median, mode, frequency counts, percentiles</li>
          </ul>
          <p>The app will automatically detect your data type or you can manually specify it.</p>
        ")
               ),
               uiOutput("dataTypeDetection"),
               conditionalPanel(
                 condition = "output.dataType == 'numerical'",
                 h4("Numerical Data Summary"),
                 div(style = "font-size: 14px;", verbatimTextOutput("numericalSummary")),
                 plotOutput("numericalPlots", height = "400px")
               ),
               conditionalPanel(
                 condition = "output.dataType == 'categorical'",
                 h4("Categorical Data Summary"),
                 div(style = "font-size: 14px;", verbatimTextOutput("categoricalSummary")),
                 plotOutput("categoricalPlots", height = "400px")
               ),
               conditionalPanel(
                 condition = "output.dataType == 'ordinal'",
                 h4("Ordinal Data Summary"),
                 div(style = "font-size: 14px;", verbatimTextOutput("ordinalSummary")),
                 plotOutput("ordinalPlots", height = "400px")
               ),
               conditionalPanel(
                 condition = "output.dataType == 'mixed'",
                 h4("Mixed Data Summary"),
                 div(style = "font-size: 14px;", verbatimTextOutput("mixedSummary"))
               ),
               conditionalPanel(
                 condition = "output.dataType == 'unknown'",
                 h4("Data Analysis"),
                 div(style = "font-size: 14px;", verbatimTextOutput("unknownData"))
               )
             )
           )
  ),
  
  # User Guide Tab
  tabPanel("User Guide",
           fluidPage(
             div(
               style = "padding: 20px; max-width: 1000px; margin: 0 auto; font-size: 16px;",
               h2("How to Use CalcuStats", style = "color: #1A5276; text-align: center;"),
               h3("Welcome to CalcuStats!"),
               p("This application provides several tools for statistical calculations including sample size determination, power analysis, and descriptive statistics."),
               
               h3("Features Overview"),
               tags$ul(
                 tags$li(strong("Proportional Allocation:"), "Allocate a specified sample size proportionally across strata."),
                 tags$li(strong("Taro Yamane:"), "Calculate sample size using the Taro Yamane formula for finite populations."),
                 tags$li(strong("Cochran Formula:"), "Calculate sample size using Cochran's formula for proportions."),
                 tags$li(strong("Other Formulas:"), "Various sample size formulas for different study designs and parameters."),
                 tags$li(strong("Power Analysis:"), "Determine required sample size for various statistical tests."),
                 tags$li(strong("Descriptive Statistics:"), "Compute basic statistics for your data.")
               ),
               
               h3("About the Developer"),
               p("CalcuStats was developed by Mudasir Mohammed Ibrahim, a Registered Nurse with a Bachelor of Science degree and Diploma qualifications."),
               p("With expertise in both healthcare and data analysis, Mudasir created this tool to help researchers and students perform essential sample size and statistical calculations with ease."),
               p("Connect with Mudasir on GitHub:", 
                 tags$a(href="https://github.com/mudassiribrahim30", target="_blank", "github.com/mudassiribrahim30")),
               
               h3("Detailed Instructions"),
               h4("Other Formulas"),
               tags$ol(
                 tags$li("Select the desired study design from the dropdown menu."),
                 tags$li("Enter the required parameters for the selected formula."),
                 tags$li("Adjust the non-response rate if needed."),
                 tags$li("Add strata for proportional allocation if needed."),
                 tags$li("The app will calculate the sample size and provide allocation."),
                 tags$li("Generate flow charts and download results as needed.")
               ),
               
               h3("Tips"),
               tags$ul(
                 tags$li("For proportional allocation, remember to account for non-response if applicable."),
                 tags$li("In the flow chart editor, click on nodes or edges to edit them."),
                 tags$li("Use the download buttons to save your results for reporting."),
                 tags$li("For large populations, Taro Yamane provides a quick estimate."),
                 tags$li("Cochran's formula is best when working with proportions.")
               ),
               
               h3("Need Help?"),
               p("For any questions or suggestions, please contact ", 
                 tags$a(href="mailto:mudassiribrahim30@gmail.com", "mudassiribrahim30@gmail.com"), ".")
             )
           )
  ),
  
  tabPanel("Usage Statistics",
           fluidPage(
             div(
               style = "text-align: center; padding: 50px; font-size: 16px;",
               h3("App Usage Statistics"),
               div(
                 style = "font-size: 24px; margin: 20px; padding: 20px; background-color: #f8f9fa; border-radius: 10px;",
                 paste("This app has been used", counter, "times.")
               ),
               p("Thank you for using CalcuStats!")
             )
           )
  ),
  
  footer = div(
    class = "who-footer",
    HTML("<p>© 2025 Mudasir Mohammed Ibrahim. All rights reserved. | 
         <a href='https://github.com/mudassiribrahim30' target='_blank'>GitHub Profile</a> | 
         <a href='mailto:mudassiribrahim30@gmail.com'>Contact Developer</a></p>"),
    div(
      style = "margin-top: 10px; font-size: 0.9em; color: #fff;",
      "Your Professional Companion for Statistical Analysis and Sample Size Calculation"
    ),
    div(
      style = "margin-top: 15px; font-size: 0.8em; color: #ccc;",
      "We welcome suggestions for new formulas and features. Contact us to contribute!"
    )
  )
)

# Server logic with all the requested improvements
server <- function(input, output, session) {
  
  # Initialize reactive values for storing inputs
  initValues <- reactiveValues(
    proportional = list(
      custom_sample = 100,
      non_response_custom = 0,
      stratumCount = 1,
      stratum_names = c("Stratum 1"),
      stratum_pops = c(100)
    ),
    yamane = list(
      e = 0.05,
      non_response = 0,
      stratumCount = 1,
      stratum_names = c("Stratum 1"),
      stratum_pops = c(100)
    ),
    cochran = list(
      p_input = "0.5",
      z = 1.96,
      e_c = 0.05,
      N_c = NULL,
      non_response_c = 0,
      stratumCount = 1,
      stratum_names = c("Stratum 1"),
      stratum_pops = c(100)
    ),
    other = list(
      formula_type = "single_proportion",
      alpha_other = 0.05,
      power_other = 0.8,
      effect_size_other = 0.5,
      sigma_other = 1,
      d_other = 0.1,
      p_other = 0.5,
      p1_other = 0.5,
      p2_other = 0.3,
      r_other = 0.3,
      beta_other = 0.2,
      r_squared_other = 0.25,
      or_other = 2.0,
      rr_other = 1.5,
      prevalence_other = 0.1,
      precision_other = 0.05,
      non_response_other = 0,
      stratumCount = 1,
      stratum_names = c("Stratum 1"),
      stratum_pops = c(100)
    ),
    power = list(
      testType = "Independent t-test",
      effectSize = 0.5,
      alpha = 0.05,
      power = 0.8,
      predictors = 2
    ),
    desc = list(
      dataInput = ""
    )
  )
  
  # Load saved values from localStorage if available
  observe({
    # Try to load values from localStorage
    tryCatch({
      # Proportional Allocation
      if (!is.null(input$proportional_values)) {
        prop_vals <- jsonlite::fromJSON(input$proportional_values)
        initValues$proportional <- prop_vals
        updateNumericInput(session, "custom_sample", value = prop_vals$custom_sample)
        updateNumericInput(session, "non_response_custom", value = prop_vals$non_response_custom)
        
        # Update strata inputs
        if (prop_vals$stratumCount > 1) {
          for (i in 1:prop_vals$stratumCount) {
            updateTextInput(session, paste0("stratum_custom", i), value = prop_vals$stratum_names[i])
            updateNumericInput(session, paste0("pop_custom", i), value = prop_vals$stratum_pops[i])
          }
        }
      }
      
      # Taro Yamane
      if (!is.null(input$yamane_values)) {
        yamane_vals <- jsonlite::fromJSON(input$yamane_values)
        initValues$yamane <- yamane_vals
        updateNumericInput(session, "e", value = yamane_vals$e)
        updateNumericInput(session, "non_response", value = yamane_vals$non_response)
        
        # Update strata inputs
        if (yamane_vals$stratumCount > 1) {
          for (i in 1:yamane_vals$stratumCount) {
            updateTextInput(session, paste0("stratum", i), value = yamane_vals$stratum_names[i])
            updateNumericInput(session, paste0("pop", i), value = yamane_vals$stratum_pops[i])
          }
        }
      }
      
      # Cochran
      if (!is.null(input$cochran_values)) {
        cochran_vals <- jsonlite::fromJSON(input$cochran_values)
        initValues$cochran <- cochran_vals
        updateTextInput(session, "p_input", value = cochran_vals$p_input)
        updateNumericInput(session, "z", value = cochran_vals$z)
        updateNumericInput(session, "e_c", value = cochran_vals$e_c)
        updateNumericInput(session, "N_c", value = cochran_vals$N_c)
        updateNumericInput(session, "non_response_c", value = cochran_vals$non_response_c)
        
        # Update strata inputs
        if (cochran_vals$stratumCount > 1) {
          for (i in 1:cochran_vals$stratumCount) {
            updateTextInput(session, paste0("stratum_c", i), value = cochran_vals$stratum_names[i])
            updateNumericInput(session, paste0("pop_c", i), value = cochran_vals$stratum_pops[i])
          }
        }
      }
      
      # Other Formulas
      if (!is.null(input$other_values)) {
        other_vals <- jsonlite::fromJSON(input$other_values)
        initValues$other <- other_vals
        updateSelectInput(session, "formula_type", selected = other_vals$formula_type)
        updateNumericInput(session, "alpha_other", value = other_vals$alpha_other)
        updateNumericInput(session, "power_other", value = other_vals$power_other)
        updateNumericInput(session, "effect_size_other", value = other_vals$effect_size_other)
        updateNumericInput(session, "sigma_other", value = other_vals$sigma_other)
        updateNumericInput(session, "d_other", value = other_vals$d_other)
        updateNumericInput(session, "p_other", value = other_vals$p_other)
        updateNumericInput(session, "p1_other", value = other_vals$p1_other)
        updateNumericInput(session, "p2_other", value = other_vals$p2_other)
        updateNumericInput(session, "r_other", value = other_vals$r_other)
        updateNumericInput(session, "beta_other", value = other_vals$beta_other)
        updateNumericInput(session, "r_squared_other", value = other_vals$r_squared_other)
        updateNumericInput(session, "or_other", value = other_vals$or_other)
        updateNumericInput(session, "rr_other", value = other_vals$rr_other)
        updateNumericInput(session, "prevalence_other", value = other_vals$prevalence_other)
        updateNumericInput(session, "precision_other", value = other_vals$precision_other)
        updateNumericInput(session, "non_response_other", value = other_vals$non_response_other)
        
        # Update strata inputs
        if (other_vals$stratumCount > 1) {
          for (i in 1:other_vals$stratumCount) {
            updateTextInput(session, paste0("stratum_other", i), value = other_vals$stratum_names[i])
            updateNumericInput(session, paste0("pop_other", i), value = other_vals$stratum_pops[i])
          }
        }
      }
      
      # Power Analysis
      if (!is.null(input$power_values)) {
        power_vals <- jsonlite::fromJSON(input$power_values)
        initValues$power <- power_vals
        updateSelectInput(session, "testType", selected = power_vals$testType)
        updateNumericInput(session, "effectSize", value = power_vals$effectSize)
        updateNumericInput(session, "alpha", value = power_vals$alpha)
        updateNumericInput(session, "power", value = power_vals$power)
        updateNumericInput(session, "predictors", value = power_vals$predictors)
      }
      
      # Descriptive Statistics
      if (!is.null(input$desc_values)) {
        desc_vals <- jsonlite::fromJSON(input$desc_values)
        initValues$desc <- desc_vals
        updateTextAreaInput(session, "dataInput", value = desc_vals$dataInput)
      }
    }, error = function(e) {
      message("Error loading saved values: ", e$message)
    })
  })
  
  # Save values to localStorage when they change
  observe({
    # Proportional Allocation
    prop_vals <- list(
      custom_sample = input$custom_sample,
      non_response_custom = input$non_response_custom,
      stratumCount = rv_custom$stratumCount,
      stratum_names = sapply(1:rv_custom$stratumCount, function(i) input[[paste0("stratum_custom", i)]]),
      stratum_pops = sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]])
    )
    session$sendCustomMessage(type = "saveValues", 
                              message = list(section = "proportional", values = jsonlite::toJSON(prop_vals)))
    
    # Taro Yamane
    yamane_vals <- list(
      e = input$e,
      non_response = input$non_response,
      stratumCount = rv$stratumCount,
      stratum_names = sapply(1:rv$stratumCount, function(i) input[[paste0("stratum", i)]]),
      stratum_pops = sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]])
    )
    session$sendCustomMessage(type = "saveValues", 
                              message = list(section = "yamane", values = jsonlite::toJSON(yamane_vals)))
    
    # Cochran
    cochran_vals <- list(
      p_input = input$p_input,
      z = input$z,
      e_c = input$e_c,
      N_c = input$N_c,
      non_response_c = input$non_response_c,
      stratumCount = rv_c$stratumCount,
      stratum_names = sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]]),
      stratum_pops = sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]])
    )
    session$sendCustomMessage(type = "saveValues", 
                              message = list(section = "cochran", values = jsonlite::toJSON(cochran_vals)))
    
    # Other Formulas
    other_vals <- list(
      formula_type = input$formula_type,
      alpha_other = input$alpha_other,
      power_other = input$power_other,
      effect_size_other = input$effect_size_other,
      sigma_other = input$sigma_other,
      d_other = input$d_other,
      p_other = input$p_other,
      p1_other = input$p1_other,
      p2_other = input$p2_other,
      r_other = input$r_other,
      beta_other = input$beta_other,
      r_squared_other = input$r_squared_other,
      or_other = input$or_other,
      rr_other = input$rr_other,
      prevalence_other = input$prevalence_other,
      precision_other = input$precision_other,
      non_response_other = input$non_response_other,
      stratumCount = rv_other$stratumCount,
      stratum_names = sapply(1:rv_other$stratumCount, function(i) input[[paste0("stratum_other", i)]]),
      stratum_pops = sapply(1:rv_other$stratumCount, function(i) input[[paste0("pop_other", i)]])
    )
    session$sendCustomMessage(type = "saveValues", 
                              message = list(section = "other", values = jsonlite::toJSON(other_vals)))
    
    # Power Analysis
    power_vals <- list(
      testType = input$testType,
      effectSize = input$effectSize,
      alpha = input$alpha,
      power = input$power,
      predictors = input$predictors
    )
    session$sendCustomMessage(type = "saveValues", 
                              message = list(section = "power", values = jsonlite::toJSON(power_vals)))
    
    # Descriptive Statistics
    desc_vals <- list(
      dataInput = input$dataInput
    )
    session$sendCustomMessage(type = "saveValues", 
                              message = list(section = "desc", values = jsonlite::toJSON(desc_vals)))
  })
  
  # Reset buttons
  observeEvent(input$reset_proportional, {
    updateNumericInput(session, "custom_sample", value = initValues$proportional$custom_sample)
    updateNumericInput(session, "non_response_custom", value = initValues$proportional$non_response_custom)
    
    # Reset strata
    rv_custom$stratumCount <- 1
    updateTextInput(session, "stratum_custom1", value = initValues$proportional$stratum_names[1])
    updateNumericInput(session, "pop_custom1", value = initValues$proportional$stratum_pops[1])
    
    # Clear saved values
    session$sendCustomMessage(type = "clearValues", message = "proportional")
  })
  
  observeEvent(input$reset_yamane, {
    updateNumericInput(session, "e", value = initValues$yamane$e)
    updateNumericInput(session, "non_response", value = initValues$yamane$non_response)
    
    # Reset strata
    rv$stratumCount <- 1
    updateTextInput(session, "stratum1", value = initValues$yamane$stratum_names[1])
    updateNumericInput(session, "pop1", value = initValues$yamane$stratum_pops[1])
    
    # Clear saved values
    session$sendCustomMessage(type = "clearValues", message = "yamane")
  })
  
  observeEvent(input$reset_cochran, {
    updateTextInput(session, "p_input", value = initValues$cochran$p_input)
    updateNumericInput(session, "z", value = initValues$cochran$z)
    updateNumericInput(session, "e_c", value = initValues$cochran$e_c)
    updateNumericInput(session, "N_c", value = initValues$cochran$N_c)
    updateNumericInput(session, "non_response_c", value = initValues$cochran$non_response_c)
    
    # Reset strata
    rv_c$stratumCount <- 1
    updateTextInput(session, "stratum_c1", value = initValues$cochran$stratum_names[1])
    updateNumericInput(session, "pop_c1", value = initValues$cochran$stratum_pops[1])
    
    # Clear saved values
    session$sendCustomMessage(type = "clearValues", message = "cochran")
  })
  
  observeEvent(input$reset_other, {
    updateSelectInput(session, "formula_type", selected = initValues$other$formula_type)
    updateNumericInput(session, "alpha_other", value = initValues$other$alpha_other)
    updateNumericInput(session, "power_other", value = initValues$other$power_other)
    updateNumericInput(session, "effect_size_other", value = initValues$other$effect_size_other)
    updateNumericInput(session, "sigma_other", value = initValues$other$sigma_other)
    updateNumericInput(session, "d_other", value = initValues$other$d_other)
    updateNumericInput(session, "p_other", value = initValues$other$p_other)
    updateNumericInput(session, "p1_other", value = initValues$other$p1_other)
    updateNumericInput(session, "p2_other", value = initValues$other$p2_other)
    updateNumericInput(session, "r_other", value = initValues$other$r_other)
    updateNumericInput(session, "beta_other", value = initValues$other$beta_other)
    updateNumericInput(session, "r_squared_other", value = initValues$other$r_squared_other)
    updateNumericInput(session, "or_other", value = initValues$other$or_other)
    updateNumericInput(session, "rr_other", value = initValues$other$rr_other)
    updateNumericInput(session, "prevalence_other", value = initValues$other$prevalence_other)
    updateNumericInput(session, "precision_other", value = initValues$other$precision_other)
    updateNumericInput(session, "non_response_other", value = initValues$other$non_response_other)
    
    # Reset strata
    rv_other$stratumCount <- 1
    updateTextInput(session, "stratum_other1", value = initValues$other$stratum_names[1])
    updateNumericInput(session, "pop_other1", value = initValues$other$stratum_pops[1])
    
    # Clear saved values
    session$sendCustomMessage(type = "clearValues", message = "other")
  })
  
  observeEvent(input$reset_power, {
    updateSelectInput(session, "testType", selected = initValues$power$testType)
    updateNumericInput(session, "effectSize", value = initValues$power$effectSize)
    updateNumericInput(session, "alpha", value = initValues$power$alpha)
    updateNumericInput(session, "power", value = initValues$power$power)
    updateNumericInput(session, "predictors", value = initValues$power$predictors)
    
    # Clear saved values
    session$sendCustomMessage(type = "clearValues", message = "power")
  })
  
  observeEvent(input$reset_desc, {
    updateTextAreaInput(session, "dataInput", value = initValues$desc$dataInput)
    
    # Clear saved values
    session$sendCustomMessage(type = "clearValues", message = "desc")
  })
  
  # Custom Proportional Allocation section variables
  rv_custom <- reactiveValues(
    stratumCount = 1,
    selectedNode = NULL,
    selectedEdge = NULL,
    nodeTexts = list(),
    edgeTexts = list(),
    nodeFontSizes = list(),
    edgeFontSizes = list()
  )
  
  observeEvent(input$addStratum_custom, { rv_custom$stratumCount <- rv_custom$stratumCount + 1 })
  observeEvent(input$removeStratum_custom, { if (rv_custom$stratumCount > 1) rv_custom$stratumCount <- rv_custom$stratumCount - 1 })
  
  output$stratumInputs_custom <- renderUI({
    lapply(1:rv_custom$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum_custom", i),
                      label = paste("Stratum", i, "Name"),
                      value = ifelse(i <= length(initValues$proportional$stratum_names), 
                                     initValues$proportional$stratum_names[i], 
                                     paste("Stratum", i)),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop_custom", i),
                         label = paste("Stratum", i, "Population"),
                         value = ifelse(i <= length(initValues$proportional$stratum_pops), 
                                        initValues$proportional$stratum_pops[i], 
                                        100),
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  output$stratumCount_custom <- reactive({
    rv_custom$stratumCount
  })
  outputOptions(output, "stratumCount_custom", suspendWhenHidden = FALSE)
  
  totalPopulation_custom <- reactive({
    sum(sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]]), na.rm = TRUE)
  })
  
  adjustedSampleSize_custom <- reactive({
    n <- input$custom_sample
    if (is.null(n) || n == 0) return(0)
    non_response_rate <- input$non_response_custom / 100
    if (non_response_rate >= 1) return(Inf) # Handle 100% non-response edge case
    ceiling(n / (1 - non_response_rate))
  })
  
  output$adjustedSampleSizeCustom <- renderText({
    n <- input$custom_sample
    adj_n <- adjustedSampleSize_custom()
    non_response_rate <- input$non_response_custom
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", n, "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", n, " / (1 - ", non_response_rate/100, ") = ", n / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  allocationData_custom <- reactive({
    N <- totalPopulation_custom()
    n <- adjustedSampleSize_custom()
    if (N == 0 || n == 0) return(NULL)
    
    strata <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("stratum_custom", i)]] )
    pops <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]] )
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv_custom$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  interpretationText_custom <- reactive({
    alloc <- allocationData_custom()
    if (is.null(alloc)) return("Please provide valid population and sample size values.")
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    paste0("Based on your specified sample size of ", input$custom_sample,
           " and a non-response rate of ", input$non_response_custom, "%",
           ", the adjusted sample size is ", adjustedSampleSize_custom(), ". ",
           paste(sentences, collapse = " "))
  })
  output$totalPopCustom <- renderText({ paste("Total Population:", totalPopulation_custom()) })
  output$sampleSizeCustom <- renderText({ paste("Your Sample Size:", input$custom_sample) })
  output$allocationTableCustom <- renderTable({ allocationData_custom() })
  output$interpretationTextCustom <- renderText({ interpretationText_custom() })
  
  generateDotCode_custom <- function() {
    if (rv_custom$stratumCount <= 1) return("")
    
    strata <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("stratum_custom", i)]] )
    pops <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]] )
    alloc <- allocationData_custom()
    total_sample <- sum(alloc$Proportional_Sample[1:rv_custom$stratumCount])
    
    nodes <- paste0("node", 1:(rv_custom$stratumCount + 3))
    
    # Create node labels based on user preferences
    node_labels <- c(paste0("Total Population\nN = ", totalPopulation_custom()))
    
    # Add stratum nodes with optional content
    for (i in 1:rv_custom$stratumCount) {
      stratum_label <- paste0("Stratum: ", strata[i])
      if (input$includePopSize_custom) {
        stratum_label <- paste0(stratum_label, "\nPop: ", pops[i])
      }
      if (input$includeStratumSize_custom) {
        stratum_label <- paste0(stratum_label, "\nSample: ", alloc$Proportional_Sample[i])
      }
      node_labels <- c(node_labels, stratum_label)
    }
    
    node_labels <- c(node_labels, paste0("Total Sample\nn = ", total_sample))
    
    for (i in seq_along(nodes)) {
      if (!is.null(rv_custom$nodeTexts[[nodes[i]]])) {
        node_labels[i] <- rv_custom$nodeTexts[[nodes[i]]]
      }
    }
    
    edge_labels <- sapply(1:rv_custom$stratumCount, function(i) {
      edge_name <- paste0("edge", i)
      if (!is.null(rv_custom$edgeTexts[[edge_name]])) {
        return(rv_custom$edgeTexts[[edge_name]])
      } else {
        return(paste0("", alloc$Proportional_Sample[i]))
      }
    })
    
    # Get font sizes for nodes and edges
    node_font_sizes <- sapply(nodes, function(n) {
      rv_custom$nodeFontSizes[[n]] %||% input$nodeFontSize_custom
    })
    
    edge_font_sizes <- sapply(paste0("edge", 1:rv_custom$stratumCount), function(e) {
      rv_custom$edgeFontSizes[[e]] %||% input$edgeFontSize_custom
    })
    
    dot_code <- paste0(
      "digraph flowchart {
        rankdir=TB;
        layout=\"", input$flowchartLayout_custom, "\";
        node [fontname=Arial, shape=\"", input$nodeShape_custom, "\", style=filled, fillcolor='", input$nodeColor_custom, "', 
              width=", input$nodeWidth_custom, ", height=", input$nodeHeight_custom, "];
        edge [color='", input$edgeColor_custom, "', arrowsize=", input$arrowSize_custom, "];
        
        // Nodes with custom font sizes
        '", nodes[1], "' [label='", node_labels[1], "', id='", nodes[1], "', fontsize=", node_font_sizes[1], "];
        '", nodes[length(nodes)], "' [label='", node_labels[length(node_labels)], "', id='", nodes[length(nodes)], "', fontsize=", node_font_sizes[length(nodes)], "];
      ",
      paste0(sapply(1:rv_custom$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' [label='", node_labels[i+1], "', id='", nodes[i+1], "', fontsize=", node_font_sizes[i+1], "];")
      }), collapse = "\n"),
      "
        
        // Edges with custom font sizes
        '", nodes[1], "' -> {",
      paste0("'", nodes[2:(rv_custom$stratumCount+1)], "'", collapse = " "),
      "} [label='', id='stratification_edges', fontsize=", input$edgeFontSize_custom, "];
      ",
      paste0(sapply(1:rv_custom$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' -> '", nodes[length(nodes)], "' [label='", edge_labels[i], "', id='edge", i, "', fontsize=", edge_font_sizes[i], "];")
      }), collapse = "\n"),
      "
      }"
    )
    
    return(dot_code)
  }
  
  output$flowchart_custom <- renderGrViz({
    if (!input$showFlowchart_custom || rv_custom$stratumCount <= 1) return()
    
    dot_code <- generateDotCode_custom()
    grViz(dot_code)
  })
  
  # JavaScript for handling clicks on nodes and edges
  jsCode_custom <- '
  $(document).ready(function() {
    document.getElementById("flowchart_custom").addEventListener("click", function(event) {
      var target = event.target;
      if (target.tagName === "text") {
        target = target.parentNode;
      }
      
      if (target.getAttribute("class") && target.getAttribute("class").includes("node")) {
        var nodeId = target.getAttribute("id");
        Shiny.setInputValue("selected_node_custom", nodeId);
      } else if (target.getAttribute("class") && target.getAttribute("class").includes("edge")) {
        var pathId = target.getAttribute("id");
        Shiny.setInputValue("selected_edge_custom", pathId);
      }
    });
  });
  '
  
  observe({
    session$sendCustomMessage(type='jsCode', list(value = jsCode_custom))
  })
  
  observeEvent(input$selected_node_custom, {
    rv_custom$selectedNode <- input$selected_node_custom
    rv_custom$selectedEdge <- NULL
    
    if (!is.null(rv_custom$nodeTexts[[rv_custom$selectedNode]])) {
      updateTextInput(session, "nodeText_custom", value = rv_custom$nodeTexts[[rv_custom$selectedNode]])
    } else {
      node_index <- as.numeric(gsub("node", "", rv_custom$selectedNode))
      if (node_index == 1) {
        updateTextInput(session, "nodeText_custom", value = paste0("Total Population\nN = ", totalPopulation_custom()))
      } else if (node_index == rv_custom$stratumCount + 2) {
        alloc <- allocationData_custom()
        total_sample <- sum(alloc$Proportional_Sample[1:rv_custom$stratumCount])
        updateTextInput(session, "nodeText_custom", 
                        value = paste0("Total Sample\nn = ", total_sample))
      } else {
        stratum_num <- node_index - 1
        strata <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("stratum_custom", i)]])
        pops <- sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]])
        alloc <- allocationData_custom()
        updateTextInput(session, "nodeText_custom", 
                        value = paste0("Stratum: ", strata[stratum_num], "\nPop: ", pops[stratum_num], "\nSample: ", alloc$Proportional_Sample[stratum_num]))
      }
    }
    
    # Update font size slider for selected node
    if (!is.null(rv_custom$nodeFontSizes[[rv_custom$selectedNode]])) {
      updateSliderInput(session, "selectedNodeFontSize_custom", 
                        value = rv_custom$nodeFontSizes[[rv_custom$selectedNode]])
    } else {
      updateSliderInput(session, "selectedNodeFontSize_custom", value = input$nodeFontSize_custom)
    }
  })
  
  observeEvent(input$selected_edge_custom, {
    rv_custom$selectedEdge <- input$selected_edge_custom
    rv_custom$selectedNode <- NULL
    
    if (!is.null(rv_custom$edgeTexts[[rv_custom$selectedEdge]])) {
      updateTextInput(session, "nodeText_custom", value = rv_custom$edgeTexts[[rv_custom$selectedEdge]])
    } else {
      edge_num <- as.numeric(gsub("edge", "", rv_custom$selectedEdge))
      alloc <- allocationData_custom()
      updateTextInput(session, "nodeText_custom", 
                      value = paste0("", alloc$Proportional_Sample[edge_num]))
    }
    
    # Update font size slider for selected edge
    if (!is.null(rv_custom$edgeFontSizes[[rv_custom$selectedEdge]])) {
      updateSliderInput(session, "selectedEdgeFontSize_custom", 
                        value = rv_custom$edgeFontSizes[[rv_custom$selectedEdge]])
    } else {
      updateSliderInput(session, "selectedEdgeFontSize_custom", value = input$edgeFontSize_custom)
    }
  })
  
  observeEvent(input$editNodeText_custom, {
    if (input$newText_custom == "") return()
    
    if (!is.null(rv_custom$selectedNode)) {
      if (grepl("->", input$newText_custom)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_custom, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_custom$nodeTexts[[rv_custom$selectedNode]] %||% input$nodeText_custom
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_custom$nodeTexts[[rv_custom$selectedNode]] <- updated_text
      } else {
        # Direct text replacement
        rv_custom$nodeTexts[[rv_custom$selectedNode]] <- input$newText_custom
      }
    } else if (!is.null(rv_custom$selectedEdge)) {
      if (grepl("->", input$newText_custom)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_custom, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_custom$edgeTexts[[rv_custom$selectedEdge]] %||% input$nodeText_custom
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_custom$edgeTexts[[rv_custom$selectedEdge]] <- updated_text
      } else {
        # Direct text replacement
        rv_custom$edgeTexts[[rv_custom$selectedEdge]] <- input$newText_custom
      }
    }
    
    updateTextInput(session, "newText_custom", value = "")
    output$flowchart_custom <- renderGrViz({
      dot_code <- generateDotCode_custom()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$deleteText_custom, {
    if (!is.null(rv_custom$selectedNode)) {
      if (grepl("->", input$newText_custom)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_custom, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_custom$nodeTexts[[rv_custom$selectedNode]] %||% input$nodeText_custom
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_custom$nodeTexts[[rv_custom$selectedNode]] <- updated_text
      } else {
        # Delete entire text
        rv_custom$nodeTexts[[rv_custom$selectedNode]] <- NULL
      }
    } else if (!is.null(rv_custom$selectedEdge)) {
      if (grepl("->", input$newText_custom)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_custom, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_custom$edgeTexts[[rv_custom$selectedEdge]] %||% input$nodeText_custom
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_custom$edgeTexts[[rv_custom$selectedEdge]] <- updated_text
      } else {
        # Delete entire text
        rv_custom$edgeTexts[[rv_custom$selectedEdge]] <- NULL
      }
    }
    
    updateTextInput(session, "newText_custom", value = "")
    output$flowchart_custom <- renderGrViz({
      dot_code <- generateDotCode_custom()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$applyFontSizes_custom, {
    if (!is.null(rv_custom$selectedNode)) {
      rv_custom$nodeFontSizes[[rv_custom$selectedNode]] <- input$selectedNodeFontSize_custom
    } else if (!is.null(rv_custom$selectedEdge)) {
      rv_custom$edgeFontSizes[[rv_custom$selectedEdge]] <- input$selectedEdgeFontSize_custom
    }
    
    output$flowchart_custom <- renderGrViz({
      dot_code <- generateDotCode_custom()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetDiagramText_custom, {
    rv_custom$nodeTexts <- list()
    rv_custom$edgeTexts <- list()
    rv_custom$nodeFontSizes <- list()
    rv_custom$edgeFontSizes <- list()
    updateTextInput(session, "nodeText_custom", value = "")
    updateTextInput(session, "newText_custom", value = "")
    output$flowchart_custom <- renderGrViz({
      dot_code <- generateDotCode_custom()
      grViz(dot_code)
    })
  })
  
  output$nodeEdgeEditorUI_custom <- renderUI({
    if (!is.null(rv_custom$selectedNode)) {
      tagList(
        p(strong("Editing Node:"), rv_custom$selectedNode),
        textInput("nodeText_custom", "Node Text:", value = rv_custom$nodeTexts[[rv_custom$selectedNode]] %||% ""),
        actionButton("editNodeText_custom", "Update Node Text", class = "btn-info"),
        actionButton("resetNodeText_custom", "Reset This Text", class = "btn-warning")
      )
    } else if (!is.null(rv_custom$selectedEdge)) {
      tagList(
        p(strong("Editing Edge:"), rv_custom$selectedEdge),
        textInput("nodeText_custom", "Edge Label:", value = rv_custom$edgeTexts[[rv_custom$selectedEdge]] %||% ""),
        actionButton("editNodeText_custom", "Update Edge Label", class = "btn-info"),
        actionButton("resetEdgeText_custom", "Reset This Text", class = "btn-warning")
      )
    } else {
      p("Click on a node or edge in the diagram to edit it.")
    }
  })
  
  observeEvent(input$resetNodeText_custom, {
    rv_custom$nodeTexts[[rv_custom$selectedNode]] <- NULL
    rv_custom$nodeFontSizes[[rv_custom$selectedNode]] <- NULL
    updateTextInput(session, "nodeText_custom", value = "")
    output$flowchart_custom <- renderGrViz({
      dot_code <- generateDotCode_custom()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetEdgeText_custom, {
    rv_custom$edgeTexts[[rv_custom$selectedEdge]] <- NULL
    rv_custom$edgeFontSizes[[rv_custom$selectedEdge]] <- NULL
    updateTextInput(session, "nodeText_custom", value = "")
    output$flowchart_custom <- renderGrViz({
      dot_code <- generateDotCode_custom()
      grViz(dot_code)
    })
  })
  
  output$downloadFlowchartPNG_custom <- downloadHandler(
    filename = function() {
      paste("proportional_allocation_flowchart_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_custom()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_png(tmp_file, file, width = 3000, height = 2000)
      unlink(tmp_file)
    }
  )
  
  output$downloadFlowchartSVG_custom <- downloadHandler(
    filename = function() {
      paste("proportional_allocation_flowchart_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_custom()
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(file)
    }
  )
  
  output$downloadFlowchartPDF_custom <- downloadHandler(
    filename = function() {
      paste("proportional_allocation_flowchart_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_custom()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_pdf(tmp_file, file)
      unlink(tmp_file)
    }
  )
  
  output$downloadCustomWord <- downloadHandler(
    filename = function() {
      paste0("proportional_allocation_results_", Sys.Date(), ".docx")
    },
    content = function(file) {
      withProgress(message = "Generating Word document", value = 0, {
        tryCatch({
          doc <- officer::read_docx()
          
          # ---- Title ----
          doc <- doc %>% 
            officer::body_add_par("Proportional Allocation Results", style = "heading 1") %>%
            officer::body_add_par(paste("Generated on:", Sys.Date()), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Input parameters ----
          incProgress(0.2, detail = "Adding input parameters")
          doc <- doc %>%
            officer::body_add_par("Input Parameters:", style = "heading 2") %>%
            officer::body_add_par(paste("Specified sample size:", as.character(input$custom_sample)), style = "Normal") %>%
            officer::body_add_par(paste("Non-response rate:", as.character(input$non_response_custom), "%"), style = "Normal") %>%
            officer::body_add_par(paste("Adjusted sample size:", as.character(adjustedSampleSize_custom())), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Stratum info ----
          incProgress(0.4, detail = "Adding stratum information")
          strata_data <- data.frame(
            Stratum = sapply(1:rv_custom$stratumCount, function(i) input[[paste0("stratum_custom", i)]]),
            Population = sapply(1:rv_custom$stratumCount, function(i) input[[paste0("pop_custom", i)]])
          )
          ft <- flextable::flextable(strata_data) %>%
            flextable::theme_box() %>%
            flextable::autofit()
          doc <- flextable::body_add_flextable(doc, ft) %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Allocation results ----
          incProgress(0.7, detail = "Adding allocation results")
          alloc_data <- allocationData_custom()
          if (!is.data.frame(alloc_data)) alloc_data <- as.data.frame(alloc_data)
          
          ft_alloc <- flextable::flextable(alloc_data) %>%
            flextable::theme_box() %>%
            flextable::autofit()
          doc <- flextable::body_add_flextable(doc, ft_alloc) %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Interpretation ----
          incProgress(0.9, detail = "Adding interpretation")
          interp <- interpretationText_custom()
          if (is.list(interp)) interp <- paste(unlist(interp), collapse = " ")
          
          doc <- doc %>%
            officer::body_add_par("Interpretation:", style = "heading 2") %>%
            officer::body_add_par(as.character(interp), style = "Normal")
          
          # ---- Save document ----
          print(doc, target = file)
          
        }, error = function(e) {
          showNotification(paste("Error generating document:", e$message), type = "error")
        })
      })
    }
  )
  
  output$downloadCustomSteps <- downloadHandler(
    filename = function() {
      paste("proportional_allocation_calculation_steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      steps <- c(
        "PROPORTIONAL ALLOCATION CALCULATION STEPS",
        "==========================================",
        paste("Date:", Sys.Date()),
        "",
        "INPUT PARAMETERS:",
        paste("Specified sample size:", input$custom_sample),
        paste("Non-response rate:", input$non_response_custom, "%"),
        "",
        "ADJUSTMENT FOR NON-RESPONSE:",
        paste("Original sample size: ", input$custom_sample),
        paste("Non-response rate: ", input$non_response_custom, "%"),
        paste("Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)"),
        paste("Calculation: ", input$custom_sample, " / (1 - ", input$non_response_custom/100, ") = ", 
              input$custom_sample / (1 - input$non_response_custom/100)),
        paste("Apply ceiling function: ", adjustedSampleSize_custom()),
        paste("Final adjusted sample size: ", adjustedSampleSize_custom()),
        "",
        "STRATUM INFORMATION:"
      )
      
      for (i in 1:rv_custom$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", input[[paste0("stratum_custom", i)]], 
                         "- Population:", input[[paste0("pop_custom", i)]]))
      }
      
      steps <- c(steps,
                 paste("Total population:", totalPopulation_custom()),
                 "",
                 "PROPORTIONAL ALLOCATION CALCULATIONS:",
                 "Formula: n_i = (N_i / N) × n"
      )
      
      alloc <- allocationData_custom()
      for (i in 1:rv_custom$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", alloc$Calculation[i]))
      }
      
      steps <- c(steps,
                 "",
                 "FINAL ALLOCATION:"
      )
      
      for (i in 1:rv_custom$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", alloc$Proportional_Sample[i], "samples"))
      }
      
      steps <- c(steps,
                 paste("Total samples:", sum(alloc$Proportional_Sample[1:rv_custom$stratumCount])),
                 "",
                 "INTERPRETATION:",
                 interpretationText_custom()
      )
      
      writeLines(steps, file)
    }
  )
  
  # Taro Yamane section variables
  rv <- reactiveValues(
    stratumCount = 1,
    selectedNode = NULL,
    selectedEdge = NULL,
    nodeTexts = list(),
    edgeTexts = list(),
    nodeFontSizes = list(),
    edgeFontSizes = list()
  )
  
  observeEvent(input$addStratum, { rv$stratumCount <- rv$stratumCount + 1 })
  observeEvent(input$removeStratum, { if (rv$stratumCount > 1) rv$stratumCount <- rv$stratumCount - 1 })
  
  output$stratumInputs <- renderUI({
    lapply(1:rv$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum", i),
                      label = paste("Stratum", i, "Name"),
                      value = ifelse(i <= length(initValues$yamane$stratum_names), 
                                     initValues$yamane$stratum_names[i], 
                                     paste("Stratum", i)),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop", i),
                         label = paste("Stratum", i, "Population"),
                         value = ifelse(i <= length(initValues$yamane$stratum_pops), 
                                        initValues$yamane$stratum_pops[i], 
                                        100),
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  output$stratumCount <- reactive({
    rv$stratumCount
  })
  outputOptions(output, "stratumCount", suspendWhenHidden = FALSE)
  
  totalPopulation <- reactive({
    sum(sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]]), na.rm = TRUE)
  })
  
  yamaneSampleSize <- reactive({
    N <- totalPopulation()
    e <- input$e
    if (N == 0 || e == 0) return(0)
    ceiling(N / (1 + N * e^2))
  })
  
  adjustedSampleSize <- reactive({
    n <- yamaneSampleSize()
    if (n == 0) return(0)
    non_response_rate <- input$non_response / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  output$formulaExplanation <- renderText({
    N <- totalPopulation()
    e <- input$e
    n <- yamaneSampleSize()
    
    paste(
      "Taro Yamane Formula Calculation:\n",
      "1. Formula: n = N / (1 + N × e²)\n",
      "2. Calculation: ", N, " / (1 + ", N, " × ", e, "²)\n",
      "3. Step-by-step: ", N, " / (1 + ", N, " × ", e^2, ") = ", N, " / (1 + ", N * e^2, ")\n",
      "4. Result: ", N, " / ", (1 + N * e^2), " = ", N / (1 + N * e^2), "\n",
      "5. Apply ceiling function: ", n, "\n",
      "\nRequired sample size: ", n
    )
  })
  
  output$adjustedSampleSize <- renderText({
    n <- yamaneSampleSize()
    adj_n <- adjustedSampleSize()
    non_response_rate <- input$non_response
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", n, "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", n, " / (1 - ", non_response_rate/100, ") = ", n / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  allocationData <- reactive({
    N <- totalPopulation()
    n <- adjustedSampleSize()
    strata <- sapply(1:rv$stratumCount, function(i) input[[paste0("stratum", i)]] )
    pops <- sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]] )
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  interpretationText <- reactive({
    alloc <- allocationData()
    if (is.null(alloc)) return("")
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    paste0("Based on the Taro Yamane formula with a margin of error of ", input$e,
           " and a non-response rate of ", input$non_response, "%",
           ", the required sample size is ", adjustedSampleSize(), ". ",
           paste(sentences, collapse = " "))
  })
  
  output$totalPop <- renderText({ paste("Total Population:", totalPopulation()) })
  output$sampleSize <- renderText({ paste("Sample Size (Yamane):", yamaneSampleSize()) })
  output$allocationTable <- renderTable({ allocationData() })
  output$interpretationText <- renderText({ interpretationText() })
  
  generateDotCode <- function() {
    if (rv$stratumCount <= 1) return("")
    
    strata <- sapply(1:rv$stratumCount, function(i) input[[paste0("stratum", i)]] )
    pops <- sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]] )
    alloc <- allocationData()
    total_sample <- sum(alloc$Proportional_Sample[1:rv$stratumCount])
    
    nodes <- paste0("node", 1:(rv$stratumCount + 3))
    
    # Create node labels based on user preferences
    node_labels <- c(paste0("Total Population\nN = ", totalPopulation()))
    
    # Add stratum nodes with optional content
    for (i in 1:rv$stratumCount) {
      stratum_label <- paste0("Stratum: ", strata[i])
      if (input$includePopSize) {
        stratum_label <- paste0(stratum_label, "\nPop: ", pops[i])
      }
      if (input$includeStratumSize) {
        stratum_label <- paste0(stratum_label, "\nSample: ", alloc$Proportional_Sample[i])
      }
      node_labels <- c(node_labels, stratum_label)
    }
    
    node_labels <- c(node_labels, paste0("Total Sample\nn = ", total_sample))
    
    for (i in seq_along(nodes)) {
      if (!is.null(rv$nodeTexts[[nodes[i]]])) {
        node_labels[i] <- rv$nodeTexts[[nodes[i]]]
      }
    }
    
    edge_labels <- sapply(1:rv$stratumCount, function(i) {
      edge_name <- paste0("edge", i)
      if (!is.null(rv$edgeTexts[[edge_name]])) {
        return(rv$edgeTexts[[edge_name]])
      } else {
        return(paste0("", alloc$Proportional_Sample[i]))
      }
    })
    
    # Get font sizes for nodes and edges
    node_font_sizes <- sapply(nodes, function(n) {
      rv$nodeFontSizes[[n]] %||% input$nodeFontSize
    })
    
    edge_font_sizes <- sapply(paste0("edge", 1:rv$stratumCount), function(e) {
      rv$edgeFontSizes[[e]] %||% input$edgeFontSize
    })
    
    dot_code <- paste0(
      "digraph flowchart {
        rankdir=TB;
        layout=\"", input$flowchartLayout, "\";
        node [fontname=Arial, shape=\"", input$nodeShape, "\", style=filled, fillcolor='", input$nodeColor, "', 
              width=", input$nodeWidth, ", height=", input$nodeHeight, "];
        edge [color='", input$edgeColor, "', arrowsize=", input$arrowSize, "];
        
        // Nodes with custom font sizes
        '", nodes[1], "' [label='", node_labels[1], "', id='", nodes[1], "', fontsize=", node_font_sizes[1], "];
        '", nodes[length(nodes)], "' [label='", node_labels[length(node_labels)], "', id='", nodes[length(nodes)], "', fontsize=", node_font_sizes[length(nodes)], "];
      ",
      paste0(sapply(1:rv$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' [label='", node_labels[i+1], "', id='", nodes[i+1], "', fontsize=", node_font_sizes[i+1], "];")
      }), collapse = "\n"),
      "
        
        // Edges with custom font sizes
        '", nodes[1], "' -> {",
      paste0("'", nodes[2:(rv$stratumCount+1)], "'", collapse = " "),
      "} [label='', id='stratification_edges', fontsize=", input$edgeFontSize, "];
      ",
      paste0(sapply(1:rv$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' -> '", nodes[length(nodes)], "' [label='", edge_labels[i], "', id='edge", i, "', fontsize=", edge_font_sizes[i], "];")
      }), collapse = "\n"),
      "
      }"
    )
    
    return(dot_code)
  }
  
  output$flowchart <- renderGrViz({
    if (!input$showFlowchart || rv$stratumCount <= 1) return()
    
    dot_code <- generateDotCode()
    grViz(dot_code)
  })
  
  # JavaScript for handling clicks on nodes and edges
  jsCode <- '
  $(document).ready(function() {
    document.getElementById("flowchart").addEventListener("click", function(event) {
      var target = event.target;
      if (target.tagName === "text") {
        target = target.parentNode;
      }
      
      if (target.getAttribute("class") && target.getAttribute("class").includes("node")) {
        var nodeId = target.getAttribute("id");
        Shiny.setInputValue("selected_node", nodeId);
      } else if (target.getAttribute("class") && target.getAttribute("class").includes("edge")) {
        var pathId = target.getAttribute("id");
        Shiny.setInputValue("selected_edge", pathId);
      }
    });
  });
  '
  
  observe({
    session$sendCustomMessage(type='jsCode', list(value = jsCode))
  })
  
  observeEvent(input$selected_node, {
    rv$selectedNode <- input$selected_node
    rv$selectedEdge <- NULL
    
    if (!is.null(rv$nodeTexts[[rv$selectedNode]])) {
      updateTextInput(session, "nodeText", value = rv$nodeTexts[[rv$selectedNode]])
    } else {
      node_index <- as.numeric(gsub("node", "", rv$selectedNode))
      if (node_index == 1) {
        updateTextInput(session, "nodeText", value = paste0("Total Population\nN = ", totalPopulation()))
      } else if (node_index == rv$stratumCount + 2) {
        alloc <- allocationData()
        total_sample <- sum(alloc$Proportional_Sample[1:rv$stratumCount])
        updateTextInput(session, "nodeText", 
                        value = paste0("Total Sample\nn = ", total_sample))
      } else {
        stratum_num <- node_index - 1
        strata <- sapply(1:rv$stratumCount, function(i) input[[paste0("stratum", i)]])
        pops <- sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]])
        alloc <- allocationData()
        updateTextInput(session, "nodeText", 
                        value = paste0("Stratum: ", strata[stratum_num], "\nPop: ", pops[stratum_num], "\nSample: ", alloc$Proportional_Sample[stratum_num]))
      }
    }
    
    # Update font size slider for selected node
    if (!is.null(rv$nodeFontSizes[[rv$selectedNode]])) {
      updateSliderInput(session, "selectedNodeFontSize", 
                        value = rv$nodeFontSizes[[rv$selectedNode]])
    } else {
      updateSliderInput(session, "selectedNodeFontSize", value = input$nodeFontSize)
    }
  })
  
  observeEvent(input$selected_edge, {
    rv$selectedEdge <- input$selected_edge
    rv$selectedNode <- NULL
    
    if (!is.null(rv$edgeTexts[[rv$selectedEdge]])) {
      updateTextInput(session, "nodeText", value = rv$edgeTexts[[rv$selectedEdge]])
    } else {
      edge_num <- as.numeric(gsub("edge", "", rv$selectedEdge))
      alloc <- allocationData()
      updateTextInput(session, "nodeText", 
                      value = paste0("", alloc$Proportional_Sample[edge_num]))
    }
    
    # Update font size slider for selected edge
    if (!is.null(rv$edgeFontSizes[[rv$selectedEdge]])) {
      updateSliderInput(session, "selectedEdgeFontSize", 
                        value = rv$edgeFontSizes[[rv$selectedEdge]])
    } else {
      updateSliderInput(session, "selectedEdgeFontSize", value = input$edgeFontSize)
    }
  })
  
  observeEvent(input$editNodeText, {
    if (input$newText == "") return()
    
    if (!is.null(rv$selectedNode)) {
      if (grepl("->", input$newText)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv$nodeTexts[[rv$selectedNode]] %||% input$nodeText
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv$nodeTexts[[rv$selectedNode]] <- updated_text
      } else {
        # Direct text replacement
        rv$nodeTexts[[rv$selectedNode]] <- input$newText
      }
    } else if (!is.null(rv$selectedEdge)) {
      if (grepl("->", input$newText)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv$edgeTexts[[rv$selectedEdge]] %||% input$nodeText
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv$edgeTexts[[rv$selectedEdge]] <- updated_text
      } else {
        # Direct text replacement
        rv$edgeTexts[[rv$selectedEdge]] <- input$newText
      }
    }
    
    updateTextInput(session, "newText", value = "")
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$deleteText, {
    if (!is.null(rv$selectedNode)) {
      if (grepl("->", input$newText)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv$nodeTexts[[rv$selectedNode]] %||% input$nodeText
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv$nodeTexts[[rv$selectedNode]] <- updated_text
      } else {
        # Delete entire text
        rv$nodeTexts[[rv$selectedNode]] <- NULL
      }
    } else if (!is.null(rv$selectedEdge)) {
      if (grepl("->", input$newText)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv$edgeTexts[[rv$selectedEdge]] %||% input$nodeText
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv$edgeTexts[[rv$selectedEdge]] <- updated_text
      } else {
        # Delete entire text
        rv$edgeTexts[[rv$selectedEdge]] <- NULL
      }
    }
    
    updateTextInput(session, "newText", value = "")
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$applyFontSizes, {
    if (!is.null(rv$selectedNode)) {
      rv$nodeFontSizes[[rv$selectedNode]] <- input$selectedNodeFontSize
    } else if (!is.null(rv$selectedEdge)) {
      rv$edgeFontSizes[[rv$selectedEdge]] <- input$selectedEdgeFontSize
    }
    
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetDiagramText, {
    rv$nodeTexts <- list()
    rv$edgeTexts <- list()
    rv$nodeFontSizes <- list()
    rv$edgeFontSizes <- list()
    updateTextInput(session, "nodeText", value = "")
    updateTextInput(session, "newText", value = "")
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode()
      grViz(dot_code)
    })
  })
  
  output$nodeEdgeEditorUI <- renderUI({
    if (!is.null(rv$selectedNode)) {
      tagList(
        p(strong("Editing Node:"), rv$selectedNode),
        textInput("nodeText", "Node Text:", value = rv$nodeTexts[[rv$selectedNode]] %||% ""),
        actionButton("editNodeText", "Update Node Text", class = "btn-info"),
        actionButton("resetNodeText", "Reset This Text", class = "btn-warning")
      )
    } else if (!is.null(rv$selectedEdge)) {
      tagList(
        p(strong("Editing Edge:"), rv$selectedEdge),
        textInput("nodeText", "Edge Label:", value = rv$edgeTexts[[rv$selectedEdge]] %||% ""),
        actionButton("editNodeText", "Update Edge Label", class = "btn-info"),
        actionButton("resetEdgeText", "Reset This Text", class = "btn-warning")
      )
    } else {
      p("Click on a node or edge in the diagram to edit it.")
    }
  })
  
  observeEvent(input$resetNodeText, {
    rv$nodeTexts[[rv$selectedNode]] <- NULL
    rv$nodeFontSizes[[rv$selectedNode]] <- NULL
    updateTextInput(session, "nodeText", value = "")
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetEdgeText, {
    rv$edgeTexts[[rv$selectedEdge]] <- NULL
    rv$edgeFontSizes[[rv$selectedEdge]] <- NULL
    updateTextInput(session, "nodeText", value = "")
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode()
      grViz(dot_code)
    })
  })
  
  output$downloadFlowchartPNG <- downloadHandler(
    filename = function() {
      paste("yamane_allocation_flowchart_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_png(tmp_file, file, width = 3000, height = 2000)
      unlink(tmp_file)
    }
  )
  
  output$downloadFlowchartSVG <- downloadHandler(
    filename = function() {
      paste("yamane_allocation_flowchart_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode()
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(file)
    }
  )
  
  output$downloadFlowchartPDF <- downloadHandler(
    filename = function() {
      paste("yamane_allocation_flowchart_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_pdf(tmp_file, file)
      unlink(tmp_file)
    }
  )
  
  output$downloadWord <- downloadHandler(
    filename = function() {
      paste0("yamane_allocation_results_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".docx")
    },
    content = function(file) {
      withProgress(message = "Generating Word document", value = 0, {
        tryCatch({
          doc <- officer::read_docx()
          
          # ---- Title ----
          doc <- doc %>% 
            officer::body_add_par("Taro Yamane Sample Size Results", style = "heading 1") %>%
            officer::body_add_par(paste("Generated on:", Sys.Date()), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Input parameters ----
          incProgress(0.2, detail = "Adding input parameters")
          doc <- doc %>%
            officer::body_add_par("Input Parameters:", style = "heading 2") %>%
            officer::body_add_par(paste("Margin of error (e):", input$e), style = "Normal") %>%
            officer::body_add_par(paste("Non-response rate:", input$non_response, "%"), style = "Normal") %>%
            officer::body_add_par(paste("Yamane sample size:", as.character(yamaneSampleSize())), style = "Normal") %>%
            officer::body_add_par(paste("Adjusted sample size:", as.character(adjustedSampleSize())), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Formula explanation ----
          incProgress(0.4, detail = "Adding formula explanation")
          doc <- doc %>%
            officer::body_add_par("Formula Calculation:", style = "heading 2") %>%
            officer::body_add_par("Taro Yamane Formula: n = N / (1 + N × e²)", style = "Normal")
          
          # ---- Stratum info ----
          incProgress(0.6, detail = "Adding stratum information")
          strata_data <- data.frame(
            Stratum = sapply(1:rv$stratumCount, function(i) input[[paste0("stratum", i)]]),
            Population = sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]])
          )
          ft <- flextable::flextable(strata_data) %>%
            flextable::theme_box() %>%
            flextable::autofit()
          doc <- flextable::body_add_flextable(doc, ft) %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Allocation results ----
          incProgress(0.8, detail = "Adding allocation results")
          alloc_data <- allocationData()
          if (!is.data.frame(alloc_data)) alloc_data <- as.data.frame(alloc_data)
          
          ft_alloc <- flextable::flextable(alloc_data) %>%
            flextable::theme_box() %>%
            flextable::autofit()
          doc <- flextable::body_add_flextable(doc, ft_alloc) %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Interpretation ----
          incProgress(0.9, detail = "Adding interpretation")
          interp <- interpretationText()
          if (is.list(interp)) interp <- paste(unlist(interp), collapse = " ")
          
          doc <- doc %>%
            officer::body_add_par("Interpretation:", style = "heading 2") %>%
            officer::body_add_par(as.character(interp), style = "Normal")
          
          # ---- Save document ----
          print(doc, target = file)
          
        }, error = function(e) {
          showNotification(paste("Error generating document:", e$message), type = "error")
        })
      })
    }
  )
  
  output$downloadSteps <- downloadHandler(
    filename = function() {
      paste("yamane_calculation_steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      steps <- c(
        "TARO YAMANE SAMPLE SIZE CALCULATION STEPS",
        "==========================================",
        paste("Date:", Sys.Date()),
        "",
        "INPUT PARAMETERS:",
        paste("Margin of error (e):", input$e),
        paste("Non-response rate:", input$non_response, "%"),
        "",
        "TARO YAMANE FORMULA:",
        "n = N / (1 + N × e²)",
        "",
        "CALCULATION:",
        paste("Total population (N):", totalPopulation()),
        paste("Margin of error (e):", input$e),
        paste("Calculation: ", totalPopulation(), " / (1 + ", totalPopulation(), " × ", input$e, "²)"),
        paste("Step-by-step: ", totalPopulation(), " / (1 + ", totalPopulation(), " × ", input$e^2, ") = ", 
              totalPopulation(), " / (1 + ", totalPopulation() * input$e^2, ")"),
        paste("Result: ", totalPopulation(), " / ", (1 + totalPopulation() * input$e^2), " = ", 
              totalPopulation() / (1 + totalPopulation() * input$e^2)),
        paste("Apply ceiling function: ", yamaneSampleSize()),
        paste("Required sample size: ", yamaneSampleSize()),
        "",
        "ADJUSTMENT FOR NON-RESPONSE:",
        paste("Original sample size: ", yamaneSampleSize()),
        paste("Non-response rate: ", input$non_response, "%"),
        paste("Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)"),
        paste("Calculation: ", yamaneSampleSize(), " / (1 - ", input$non_response/100, ") = ", 
              yamaneSampleSize() / (1 - input$non_response/100)),
        paste("Apply ceiling function: ", adjustedSampleSize()),
        paste("Final adjusted sample size: ", adjustedSampleSize()),
        "",
        "STRATUM INFORMATION:"
      )
      
      for (i in 1:rv$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", input[[paste0("stratum", i)]], 
                         "- Population:", input[[paste0("pop", i)]]))
      }
      
      steps <- c(steps,
                 paste("Total population:", totalPopulation()),
                 "",
                 "PROPORTIONAL ALLOCATION CALCULATIONS:",
                 "Formula: n_i = (N_i / N) × n"
      )
      
      alloc <- allocationData()
      for (i in 1:rv$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", alloc$Calculation[i]))
      }
      
      steps <- c(steps,
                 "",
                 "FINAL ALLOCATION:"
      )
      
      for (i in 1:rv$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", alloc$Proportional_Sample[i], "samples"))
      }
      
      steps <- c(steps,
                 paste("Total samples:", sum(alloc$Proportional_Sample[1:rv$stratumCount])),
                 "",
                 "INTERPRETATION:",
                 interpretationText()
      )
      
      writeLines(steps, file)
    }
  )
  
  # Cochran section variables
  rv_c <- reactiveValues(
    stratumCount = 1,
    selectedNode = NULL,
    selectedEdge = NULL,
    nodeTexts = list(),
    edgeTexts = list(),
    nodeFontSizes = list(),
    edgeFontSizes = list()
  )
  
  observeEvent(input$addStratum_c, { 
    rv_c$stratumCount <- rv_c$stratumCount + 1 
  })
  
  observeEvent(input$removeStratum_c, { 
    if (rv_c$stratumCount > 1) rv_c$stratumCount <- rv_c$stratumCount - 1 
  })
  
  output$stratumInputs_c <- renderUI({
    strata_list <- lapply(1:rv_c$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum_c", i),
                      label = paste("Stratum", i, "Name"),
                      value = ifelse(i <= length(initValues$cochran$stratum_names), 
                                     initValues$cochran$stratum_names[i], 
                                     paste("Stratum", i)),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width: 100%;",
            numericInput(paste0("pop_c", i),
                         label = paste("Stratum", i, "Population"),
                         value = ifelse(i <= length(initValues$cochran$stratum_pops), 
                                        initValues$cochran$stratum_pops[i], 
                                        100),
                         min = 0,
                         width = "100%")
        )
      )
    })
    
    # Add horizontal rule only between strata, not after the last one
    if (rv_c$stratumCount > 1) {
      for (i in 1:(rv_c$stratumCount - 1)) {
        strata_list[[i]] <- tagList(strata_list[[i]], tags$hr())
      }
    }
    
    return(strata_list)
  })
  
  output$stratumCount_c <- reactive({
    rv_c$stratumCount
  })
  outputOptions(output, "stratumCount_c", suspendWhenHidden = FALSE)
  
  totalPopulation_c <- reactive({
    sum(sapply(1:rv_c$stratumCount, function(i) {
      pop_val <- input[[paste0("pop_c", i)]]
      if (is.null(pop_val) || is.na(pop_val)) 0 else pop_val
    }), na.rm = TRUE)
  })
  
  # UPDATED: Function to parse proportion input (handles both proportion and percentage)
  parseProportionInput <- function(input_str) {
    if (is.null(input_str) || input_str == "") return(0.5)
    
    # Remove any whitespace
    input_str <- trimws(input_str)
    
    # Check if input contains percentage sign
    if (grepl("%$", input_str)) {
      # Remove percentage sign and convert to proportion
      num_str <- gsub("%", "", input_str)
      num_val <- as.numeric(num_str)
      if (is.na(num_val)) return(0.5)
      return(num_val / 100)
    } else {
      # Try to parse as number
      num_val <- as.numeric(input_str)
      if (is.na(num_val)) return(0.5)
      
      # If number is between 0 and 1, assume it's already a proportion
      if (num_val >= 0 && num_val <= 1) {
        return(num_val)
      }
      # If number is between 1 and 100, assume it's a percentage
      else if (num_val > 1 && num_val <= 100) {
        return(num_val / 100)
      }
      # Otherwise return default
      else {
        return(0.5)
      }
    }
  }
  
  # Reactive version for Cochran
  parseProportionInput_reactive <- reactive({
    parseProportionInput(input$p_input)
  })
  
  # Reactive version for Other Formulas
  parseProportionInput_other <- reactive({
    if (input$formula_type == "single_proportion") {
      parseProportionInput(as.character(input$p_other))
    } else if (input$formula_type %in% c("proportion_diff", "odds_ratio", "relative_risk", 
                                         "case_control", "cohort", "prevalence")) {
      # Handle different proportion inputs for various formulas
      if (!is.null(input$p1_other) && !is.na(input$p1_other)) {
        return(input$p1_other)
      } else {
        return(0.5)
      }
    } else {
      return(0.5)
    }
  })
  
  cochranSampleSize <- reactive({
    p <- parseProportionInput_reactive()
    z <- input$z
    e <- input$e_c
    N_c <- input$N_c
    
    if (is.null(p) || is.null(z) || is.null(e) || e == 0) return(0)
    
    # Validate p is between 0 and 1
    if (p <= 0 || p >= 1) return(0)
    
    # Infinite population formula
    n0 <- (z^2 * p * (1 - p)) / (e^2)
    
    # Finite population correction
    if (!is.null(N_c) && !is.na(N_c) && N_c > 0 && n0 > 0) {
      n0 / (1 + (n0 - 1) / N_c)
    } else {
      n0
    }
  })
  
  adjustedSampleSize_c <- reactive({
    n <- cochranSampleSize()
    if (is.null(n) || n == 0) return(0)
    non_response_rate <- input$non_response_c / 100
    if (non_response_rate >= 1) return(Inf) # Handle 100% non-response edge case
    ceiling(n / (1 - non_response_rate))
  })
  
  output$cochranFormulaExplanation <- renderText({
    p <- parseProportionInput_reactive()
    z <- input$z
    e <- input$e_c
    N_c <- input$N_c
    n0 <- (z^2 * p * (1 - p)) / (e^2)
    
    # Get the original input string for display
    original_input <- input$p_input
    
    explanation <- paste(
      "Cochran's Formula Calculation:\n",
      paste("Input provided: '", original_input, "'", sep = ""),
      paste("Interpreted as proportion: ", round(p, 4), " (", round(p*100, 2), "%)", sep = ""),
      "\n1. Infinite population formula: n₀ = (Z² × p × (1-p)) / e²\n",
      "2. Calculation: (", z, "² × ", round(p, 4), " × (1-", round(p, 4), ")) / ", e, "²\n",
      "3. Step-by-step: (", z^2, " × ", round(p, 4), " × ", round((1-p), 4), ") / ", e^2, "\n",
      "4. Result: (", z^2 * p * (1-p), ") / ", e^2, " = ", n0
    )
    
    if (!is.null(N_c) && !is.na(N_c) && N_c > 0) {
      n_final <- n0 / (1 + (n0 - 1) / N_c)
      explanation <- paste(
        explanation,
        "\n\nFinite Population Correction:\n",
        "5. Formula: n = n₀ / (1 + (n₀ - 1)/N)\n",
        "6. Calculation: ", n0, " / (1 + (", n0, " - 1)/", N_c, ")\n",
        "7. Step-by-step: ", n0, " / (1 + ", (n0 - 1), "/", N_c, ")\n",
        "8. Result: ", n0, " / (1 + ", (n0 - 1)/N_c, ") = ", n0, " / ", (1 + (n0 - 1)/N_c), " = ", n_final
      )
    } else {
      explanation <- paste(explanation, "\n\nNo finite population correction applied.")
    }
    
    paste(explanation, "\n\nRequired sample size: ", ceiling(cochranSampleSize()))
  })
  
  output$cochranSample <- renderText({
    paste("Cochran Sample Size:", ceiling(cochranSampleSize()))
  })
  
  output$cochranAdjustedSample <- renderText({
    n <- ceiling(cochranSampleSize())
    adj_n <- adjustedSampleSize_c()
    non_response_rate <- input$non_response_c
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", n, "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", n, " / (1 - ", non_response_rate/100, ") = ", n / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  allocationData_c <- reactive({
    N <- totalPopulation_c()
    n <- adjustedSampleSize_c()
    if (N == 0 || n == 0 || is.null(N) || is.null(n)) return(NULL)
    
    strata <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]] )
    pops <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]] )
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv_c$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  cochranInterpretation <- reactive({
    alloc <- allocationData_c()
    if (is.null(alloc)) return("Please provide valid population and sample size values.")
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    paste0("Based on Cochran's formula with p = ", round(parseProportionInput_reactive(), 4), " (", round(parseProportionInput_reactive()*100, 2), "%)", 
           ", Z = ", input$z,
           ", e = ", input$e_c, ifelse(!is.null(input$N_c) && !is.na(input$N_c) && input$N_c > 0, 
                                       paste0(", N = ", input$N_c), ""),
           " and a non-response rate of ", input$non_response_c, "%",
           ", the required sample size is ", adjustedSampleSize_c(), ". ",
           paste(sentences, collapse = " "))
  })
  
  output$cochranAllocationTable <- renderTable({ allocationData_c() })
  output$cochranInterpretation <- renderText({ cochranInterpretation() })
  
  generateDotCode_c <- function() {
    if (rv_c$stratumCount <= 1) return("")
    
    strata <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]] )
    pops <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]] )
    alloc <- allocationData_c()
    total_sample <- sum(alloc$Proportional_Sample[1:rv_c$stratumCount])
    
    nodes <- paste0("node", 1:(rv_c$stratumCount + 3))
    
    # Create node labels based on user preferences
    node_labels <- c(paste0("Total Population\nN = ", totalPopulation_c()))
    
    # Add stratum nodes with optional content
    for (i in 1:rv_c$stratumCount) {
      stratum_label <- paste0("Stratum: ", strata[i])
      if (input$includePopSize_c) {
        stratum_label <- paste0(stratum_label, "\nPop: ", pops[i])
      }
      if (input$includeStratumSize_c) {
        stratum_label <- paste0(stratum_label, "\nSample: ", alloc$Proportional_Sample[i])
      }
      node_labels <- c(node_labels, stratum_label)
    }
    
    node_labels <- c(node_labels, paste0("Total Sample\nn = ", total_sample))
    
    for (i in seq_along(nodes)) {
      if (!is.null(rv_c$nodeTexts[[nodes[i]]])) {
        node_labels[i] <- rv_c$nodeTexts[[nodes[i]]]
      }
    }
    
    edge_labels <- sapply(1:rv_c$stratumCount, function(i) {
      edge_name <- paste0("edge", i)
      if (!is.null(rv_c$edgeTexts[[edge_name]])) {
        return(rv_c$edgeTexts[[edge_name]])
      } else {
        return(paste0("", alloc$Proportional_Sample[i]))
      }
    })
    
    # Get font sizes for nodes and edges
    node_font_sizes <- sapply(nodes, function(n) {
      rv_c$nodeFontSizes[[n]] %||% input$nodeFontSize_c
    })
    
    edge_font_sizes <- sapply(paste0("edge", 1:rv_c$stratumCount), function(e) {
      rv_c$edgeFontSizes[[e]] %||% input$edgeFontSize_c
    })
    
    dot_code <- paste0(
      "digraph flowchart {
      rankdir=TB;
      layout=\"", input$flowchartLayout_c, "\";
      node [fontname=Arial, shape=\"", input$nodeShape_c, "\", style=filled, fillcolor='", input$nodeColor_c, "', 
            width=", input$nodeWidth_c, ", height=", input$nodeHeight_c, "];
      edge [color='", input$edgeColor_c, "', arrowsize=", input$arrowSize_c, "];
      
      // Nodes with custom font sizes
      '", nodes[1], "' [label='", node_labels[1], "', id='", nodes[1], "', fontsize=", node_font_sizes[1], "];
      '", nodes[length(nodes)], "' [label='", node_labels[length(node_labels)], "', id='", nodes[length(nodes)], "', fontsize=", node_font_sizes[length(nodes)], "];
    ",
      paste0(sapply(1:rv_c$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' [label='", node_labels[i+1], "', id='", nodes[i+1], "', fontsize=", node_font_sizes[i+1], "];")
      }), collapse = "\n"),
      "
      
      // Edges with custom font sizes
      '", nodes[1], "' -> {",
      paste0("'", nodes[2:(rv_c$stratumCount+1)], "'", collapse = " "),
      "} [label='', id='stratification_edges', fontsize=", input$edgeFontSize_c, "];
    ",
      paste0(sapply(1:rv_c$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' -> '", nodes[length(nodes)], "' [label='", edge_labels[i], "', id='edge", i, "', fontsize=", edge_font_sizes[i], "];")
      }), collapse = "\n"),
      "
    }"
    )
    
    return(dot_code)
  }
  
  output$flowchart_c <- renderGrViz({
    if (!input$showFlowchart_c || rv_c$stratumCount <= 1) return()
    
    dot_code <- generateDotCode_c()
    grViz(dot_code)
  })
  
  # JavaScript for handling clicks on nodes and edges
  jsCode_c <- '
$(document).ready(function() {
  document.getElementById("flowchart_c").addEventListener("click", function(event) {
    var target = event.target;
    if (target.tagName === "text") {
      target = target.parentNode;
    }
    
    if (target.getAttribute("class") && target.getAttribute("class").includes("node")) {
      var nodeId = target.getAttribute("id");
      Shiny.setInputValue("selected_node_c", nodeId);
    } else if (target.getAttribute("class") && target.getAttribute("class").includes("edge")) {
      var pathId = target.getAttribute("id");
      Shiny.setInputValue("selected_edge_c", pathId);
    }
  });
});
'
  
  observe({
    session$sendCustomMessage(type='jsCode', list(value = jsCode_c))
  })
  
  observeEvent(input$selected_node_c, {
    rv_c$selectedNode <- input$selected_node_c
    rv_c$selectedEdge <- NULL
    
    if (!is.null(rv_c$nodeTexts[[rv_c$selectedNode]])) {
      updateTextInput(session, "nodeText_c", value = rv_c$nodeTexts[[rv_c$selectedNode]])
    } else {
      node_index <- as.numeric(gsub("node", "", rv_c$selectedNode))
      if (node_index == 1) {
        updateTextInput(session, "nodeText_c", value = paste0("Total Population\nN = ", totalPopulation_c()))
      } else if (node_index == rv_c$stratumCount + 2) {
        alloc <- allocationData_c()
        total_sample <- sum(alloc$Proportional_Sample[1:rv_c$stratumCount])
        updateTextInput(session, "nodeText_c", 
                        value = paste0("Total Sample\nn = ", total_sample))
      } else {
        stratum_num <- node_index - 1
        strata <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]])
        pops <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]])
        alloc <- allocationData_c()
        updateTextInput(session, "nodeText_c", 
                        value = paste0("Stratum: ", strata[stratum_num], "\nPop: ", pops[stratum_num], "\nSample: ", alloc$Proportional_Sample[stratum_num]))
      }
    }
    
    # Update font size slider for selected node
    if (!is.null(rv_c$nodeFontSizes[[rv_c$selectedNode]])) {
      updateSliderInput(session, "selectedNodeFontSize_c", 
                        value = rv_c$nodeFontSizes[[rv_c$selectedNode]])
    } else {
      updateSliderInput(session, "selectedNodeFontSize_c", value = input$nodeFontSize_c)
    }
  })
  
  observeEvent(input$selected_edge_c, {
    rv_c$selectedEdge <- input$selected_edge_c
    rv_c$selectedNode <- NULL
    
    if (!is.null(rv_c$edgeTexts[[rv_c$selectedEdge]])) {
      updateTextInput(session, "nodeText_c", value = rv_c$edgeTexts[[rv_c$selectedEdge]])
    } else {
      edge_num <- as.numeric(gsub("edge", "", rv_c$selectedEdge))
      alloc <- allocationData_c()
      updateTextInput(session, "nodeText_c", 
                      value = paste0("", alloc$Proportional_Sample[edge_num]))
    }
    
    # Update font size slider for selected edge
    if (!is.null(rv_c$edgeFontSizes[[rv_c$selectedEdge]])) {
      updateSliderInput(session, "selectedEdgeFontSize_c", 
                        value = rv_c$edgeFontSizes[[rv_c$selectedEdge]])
    } else {
      updateSliderInput(session, "selectedEdgeFontSize_c", value = input$edgeFontSize_c)
    }
  })
  
  observeEvent(input$editNodeText_c, {
    if (input$newText_c == "") return()
    
    if (!is.null(rv_c$selectedNode)) {
      if (grepl("->", input$newText_c)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_c, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_c$nodeTexts[[rv_c$selectedNode]] %||% input$nodeText_c
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_c$nodeTexts[[rv_c$selectedNode]] <- updated_text
      } else {
        # Direct text replacement
        rv_c$nodeTexts[[rv_c$selectedNode]] <- input$newText_c
      }
    } else if (!is.null(rv_c$selectedEdge)) {
      if (grepl("->", input$newText_c)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_c, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_c$edgeTexts[[rv_c$selectedEdge]] %||% input$nodeText_c
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_c$edgeTexts[[rv_c$selectedEdge]] <- updated_text
      } else {
        # Direct text replacement
        rv_c$edgeTexts[[rv_c$selectedEdge]] <- input$newText_c
      }
    }
    
    updateTextInput(session, "newText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_c()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$deleteText_c, {
    if (!is.null(rv_c$selectedNode)) {
      if (grepl("->", input$newText_c)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_c, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_c$nodeTexts[[rv_c$selectedNode]] %||% input$nodeText_c
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_c$nodeTexts[[rv_c$selectedNode]] <- updated_text
      } else {
        # Delete entire text
        rv_c$nodeTexts[[rv_c$selectedNode]] <- NULL
      }
    } else if (!is.null(rv_c$selectedEdge)) {
      if (grepl("->", input$newText_c)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_c, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_c$edgeTexts[[rv_c$selectedEdge]] %||% input$nodeText_c
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_c$edgeTexts[[rv_c$selectedEdge]] <- updated_text
      } else {
        # Delete entire text
        rv_c$edgeTexts[[rv_c$selectedEdge]] <- NULL
      }
    }
    
    updateTextInput(session, "newText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_c()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$applyFontSizes_c, {
    if (!is.null(rv_c$selectedNode)) {
      rv_c$nodeFontSizes[[rv_c$selectedNode]] <- input$selectedNodeFontSize_c
    } else if (!is.null(rv_c$selectedEdge)) {
      rv_c$edgeFontSizes[[rv_c$selectedEdge]] <- input$selectedEdgeFontSize_c
    }
    
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_c()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetDiagramText_c, {
    rv_c$nodeTexts <- list()
    rv_c$edgeTexts <- list()
    rv_c$nodeFontSizes <- list()
    rv_c$edgeFontSizes <- list()
    updateTextInput(session, "nodeText_c", value = "")
    updateTextInput(session, "newText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_c()
      grViz(dot_code)
    })
  })
  
  output$nodeEdgeEditorUI_c <- renderUI({
    if (!is.null(rv_c$selectedNode)) {
      tagList(
        p(strong("Editing Node:"), rv_c$selectedNode),
        textInput("nodeText_c", "Node Text:", value = rv_c$nodeTexts[[rv_c$selectedNode]] %||% ""),
        actionButton("editNodeText_c", "Update Node Text", class = "btn-info"),
        actionButton("resetNodeText_c", "Reset This Text", class = "btn-warning")
      )
    } else if (!is.null(rv_c$selectedEdge)) {
      tagList(
        p(strong("Editing Edge:"), rv_c$selectedEdge),
        textInput("nodeText_c", "Edge Label:", value = rv_c$edgeTexts[[rv_c$selectedEdge]] %||% ""),
        actionButton("editNodeText_c", "Update Edge Label", class = "btn-info"),
        actionButton("resetEdgeText_c", "Reset This Text", class = "btn-warning")
      )
    } else {
      p("Click on a node or edge in the diagram to edit it.")
    }
  })
  
  observeEvent(input$resetNodeText_c, {
    rv_c$nodeTexts[[rv_c$selectedNode]] <- NULL
    rv_c$nodeFontSizes[[rv_c$selectedNode]] <- NULL
    updateTextInput(session, "nodeText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_c()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetEdgeText_c, {
    rv_c$edgeTexts[[rv_c$selectedEdge]] <- NULL
    rv_c$edgeFontSizes[[rv_c$selectedEdge]] <- NULL
    updateTextInput(session, "nodeText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_c()
      grViz(dot_code)
    })
  })
  
  output$downloadFlowchartPNG_c <- downloadHandler(
    filename = function() {
      paste("cochran_allocation_flowchart_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_c()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_png(tmp_file, file, width = 3000, height = 2000)
      unlink(tmp_file)
    }
  )
  
  output$downloadFlowchartSVG_c <- downloadHandler(
    filename = function() {
      paste("cochran_allocation_flowchart_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_c()
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(file)
    }
  )
  
  output$downloadFlowchartPDF_c <- downloadHandler(
    filename = function() {
      paste("cochran_allocation_flowchart_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_c()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_pdf(tmp_file, file)
      unlink(tmp_file)
    }
  )
  
  output$downloadCochranWord <- downloadHandler(
    filename = function() {
      paste0("cochran_allocation_results_", Sys.Date(), ".docx")
    },
    content = function(file) {
      withProgress(message = "Generating Word document", value = 0, {
        tryCatch({
          doc <- officer::read_docx()
          
          # ---- Title ----
          doc <- doc %>% 
            officer::body_add_par("Cochran Sample Size Results", style = "heading 1") %>%
            officer::body_add_par(paste("Generated on:", Sys.Date()), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Input parameters ----
          incProgress(0.2, detail = "Adding input parameters")
          doc <- doc %>%
            officer::body_add_par("Input Parameters:", style = "heading 2") %>%
            officer::body_add_par(paste("Estimated proportion input: '", input$p_input, "'"), style = "Normal") %>%
            officer::body_add_par(paste("Interpreted proportion (p):", round(parseProportionInput_reactive(), 4), 
                                        "(", round(parseProportionInput_reactive()*100, 2), "%)"), style = "Normal") %>%
            officer::body_add_par(paste("Z-score (Z):", input$z), style = "Normal") %>%
            officer::body_add_par(paste("Margin of error (e):", input$e_c), style = "Normal")
          
          if (!is.null(input$N_c) && !is.na(input$N_c) && input$N_c > 0) {
            doc <- doc %>% officer::body_add_par(paste("Population size (N):", input$N_c), style = "Normal")
          }
          
          doc <- doc %>%
            officer::body_add_par(paste("Non-response rate:", input$non_response_c, "%"), style = "Normal") %>%
            officer::body_add_par(paste("Cochran sample size:", ceiling(cochranSampleSize())), style = "Normal") %>%
            officer::body_add_par(paste("Adjusted sample size:", adjustedSampleSize_c()), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # ---- Formula explanation ----
          incProgress(0.4, detail = "Adding formula explanation")
          doc <- doc %>%
            officer::body_add_par("Formula Calculation:", style = "heading 2") %>%
            officer::body_add_par("Cochran's Formula: n₀ = (Z² × p × (1-p)) / e²", style = "Normal")
          
          # ---- Stratum info ----
          incProgress(0.6, detail = "Adding stratum information")
          if (rv_c$stratumCount > 0) {
            doc <- doc %>% officer::body_add_par("Stratum Information:", style = "heading 2")
            
            strata_data <- data.frame(
              Stratum = sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]]),
              Population = sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]]),
              stringsAsFactors = FALSE
            )
            
            ft <- flextable::flextable(strata_data) %>%
              flextable::theme_box() %>%
              flextable::autofit()
            doc <- flextable::body_add_flextable(doc, ft) %>%
              officer::body_add_par("", style = "Normal")
          }
          
          # ---- Allocation results ----
          incProgress(0.8, detail = "Adding allocation results")
          alloc_data <- allocationData_c()
          if (!is.null(alloc_data) && nrow(alloc_data) > 0) {
            doc <- doc %>%
              officer::body_add_par("Allocation Results:", style = "heading 2")
            
            ft_alloc <- flextable::flextable(alloc_data) %>%
              flextable::theme_box() %>%
              flextable::autofit()
            doc <- flextable::body_add_flextable(doc, ft_alloc) %>%
              officer::body_add_par("", style = "Normal")
          }
          
          # ---- Interpretation ----
          incProgress(0.9, detail = "Adding interpretation")
          interp <- cochranInterpretation()
          if (is.list(interp)) interp <- paste(unlist(interp), collapse = " ")
          
          doc <- doc %>%
            officer::body_add_par("Interpretation:", style = "heading 2") %>%
            officer::body_add_par(as.character(interp), style = "Normal")
          
          # ---- Save document ----
          print(doc, target = file)
          
        }, error = function(e) {
          showNotification(paste("Error generating document:", e$message), type = "error")
        })
      })
    }
  )
  
  output$downloadCochranSteps <- downloadHandler(
    filename = function() {
      paste("cochran_calculation_steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      steps <- c(
        "COCHRAN SAMPLE SIZE CALCULATION STEPS",
        "======================================",
        paste("Date:", Sys.Date()),
        "",
        "INPUT PARAMETERS:",
        paste("Estimated proportion input: '", input$p_input, "'"),
        paste("Interpreted as proportion (p):", round(parseProportionInput_reactive(), 4), 
              "(", round(parseProportionInput_reactive()*100, 2), "%)"),
        paste("Z-score (Z):", input$z),
        paste("Margin of error (e):", input$e_c)
      )
      
      if (!is.null(input$N_c) && !is.na(input$N_c) && input$N_c > 0) {
        steps <- c(steps, paste("Population size (N):", input$N_c))
      }
      
      steps <- c(steps,
                 paste("Non-response rate:", input$non_response_c, "%"),
                 "",
                 "COCHRAN'S FORMULA:",
                 "n₀ = (Z² × p × (1-p)) / e²",
                 "",
                 "CALCULATION:",
                 paste("Z² = ", input$z, "² = ", input$z^2),
                 paste("p × (1-p) = ", round(parseProportionInput_reactive(), 4), " × (1-", round(parseProportionInput_reactive(), 4), 
                       ") = ", round(parseProportionInput_reactive() * (1-parseProportionInput_reactive()), 4)),
                 paste("Z² × p × (1-p) = ", input$z^2, " × ", 
                       round(parseProportionInput_reactive() * (1-parseProportionInput_reactive()), 4), 
                       " = ", round(input$z^2 * parseProportionInput_reactive() * (1-parseProportionInput_reactive()), 4)),
                 paste("e² = ", input$e_c, "² = ", input$e_c^2),
                 paste("n₀ = ", round(input$z^2 * parseProportionInput_reactive() * (1-parseProportionInput_reactive()), 4), 
                       " / ", input$e_c^2, " = ", 
                       round((input$z^2 * parseProportionInput_reactive() * (1-parseProportionInput_reactive())) / (input$e_c^2), 4))
      )
      
      if (!is.null(input$N_c) && !is.na(input$N_c) && input$N_c > 0) {
        n0 <- (input$z^2 * parseProportionInput_reactive() * (1-parseProportionInput_reactive())) / (input$e_c^2)
        steps <- c(steps,
                   "",
                   "FINITE POPULATION CORRECTION:",
                   "n = n₀ / (1 + (n₀ - 1)/N)",
                   paste("Calculation: ", round(n0, 4), " / (1 + (", round(n0, 4), " - 1)/", input$N_c, ")"),
                   paste("Step-by-step: ", round(n0, 4), " / (1 + ", round(n0 - 1, 4), "/", input$N_c, ") = ", 
                         round(n0, 4), " / (1 + ", round((n0 - 1)/input$N_c, 4), ")"),
                   paste("Result: ", round(n0, 4), " / ", round((1 + (n0 - 1)/input$N_c), 4), " = ", 
                         round(n0 / (1 + (n0 - 1)/input$N_c), 4))
        )
      }
      
      steps <- c(steps,
                 paste("Apply ceiling function: ", ceiling(cochranSampleSize())),
                 paste("Required sample size: ", ceiling(cochranSampleSize())),
                 "",
                 "ADJUSTMENT FOR NON-RESPONSE:",
                 paste("Original sample size: ", ceiling(cochranSampleSize())),
                 paste("Non-response rate: ", input$non_response_c, "%"),
                 paste("Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)"),
                 paste("Calculation: ", ceiling(cochranSampleSize()), " / (1 - ", input$non_response_c/100, ") = ", 
                       ceiling(cochranSampleSize()) / (1 - input$non_response_c/100)),
                 paste("Apply ceiling function: ", adjustedSampleSize_c()),
                 paste("Final adjusted sample size: ", adjustedSampleSize_c()),
                 "",
                 "STRATUM INFORMATION:"
      )
      
      for (i in 1:rv_c$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", input[[paste0("stratum_c", i)]], 
                         "- Population:", input[[paste0("pop_c", i)]]))
      }
      
      steps <- c(steps,
                 paste("Total population:", totalPopulation_c()),
                 "",
                 "PROPORTIONAL ALLOCATION CALCULATIONS:",
                 "Formula: n_i = (N_i / N) × n"
      )
      
      alloc <- allocationData_c()
      for (i in 1:rv_c$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", alloc$Calculation[i]))
      }
      
      steps <- c(steps,
                 "",
                 "FINAL ALLOCATION:"
      )
      
      for (i in 1:rv_c$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", alloc$Proportional_Sample[i], "samples"))
      }
      
      steps <- c(steps,
                 paste("Total samples:", sum(alloc$Proportional_Sample[1:rv_c$stratumCount])),
                 "",
                 "INTERPRETATION:",
                 cochranInterpretation()
      )
      
      writeLines(steps, file)
    }
  )
  
  # Other Formulas section variables
  rv_other <- reactiveValues(
    stratumCount = 1,
    selectedNode = NULL,
    selectedEdge = NULL,
    nodeTexts = list(),
    edgeTexts = list(),
    nodeFontSizes = list(),
    edgeFontSizes = list()
  )
  
  observeEvent(input$addStratum_other, { rv_other$stratumCount <- rv_other$stratumCount + 1 })
  observeEvent(input$removeStratum_other, { if (rv_other$stratumCount > 1) rv_other$stratumCount <- rv_other$stratumCount - 1 })
  
  output$stratumInputs_other <- renderUI({
    lapply(1:rv_other$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum_other", i),
                      label = paste("Stratum", i, "Name"),
                      value = ifelse(i <= length(initValues$other$stratum_names), 
                                     initValues$other$stratum_names[i], 
                                     paste("Stratum", i)),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop_other", i),
                         label = paste("Stratum", i, "Population"),
                         value = ifelse(i <= length(initValues$other$stratum_pops), 
                                        initValues$other$stratum_pops[i], 
                                        100),
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  output$stratumCount_other <- reactive({
    rv_other$stratumCount
  })
  outputOptions(output, "stratumCount_other", suspendWhenHidden = FALSE)
  
  totalPopulation_other <- reactive({
    sum(sapply(1:rv_other$stratumCount, function(i) input[[paste0("pop_other", i)]]), na.rm = TRUE)
  })
  
  # UPDATED: Formula parameters with Single Population Proportion percentage support
  output$formula_params <- renderUI({
    formula_type <- input$formula_type
    
    params <- switch(formula_type,
                     "single_proportion" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("p_other", "Estimated Proportion (p) or Percentage (%)", 
                                 value = "0.5",
                                 placeholder = "Enter as proportion (0.5) or percentage (50%)"),
                       div(
                         class = "who-info-box",
                         style = "margin-top: 5px; font-size: 12px;",
                         HTML("<strong>Flexible Input Format:</strong><br>
                              • Enter as proportion: <em>0.5</em> (for 50%)<br>
                              • Enter as percentage: <em>50%</em> or <em>50</em><br>
                              • Both will give same results<br>
                              <em>Examples: 0.3, 30%, 5%, 0.05</em>")
                       ),
                       numericInput("d_other", "Margin of Error (d)", value = 0.05, min = 0.01, step = 0.01)
                     ),
                     "mean_known_var" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("sigma_other", "Standard Deviation (σ)", value = 1, min = 0.01, step = 0.1),
                       numericInput("d_other", "Margin of Error (d)", value = 0.1, min = 0.01, step = 0.01)
                     ),
                     "mean_unknown_var" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("effect_size_other", "Effect Size (Cohen's d)", value = 0.5, min = 0.1, step = 0.1),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01)
                     ),
                     "proportion_diff" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("p1_other", "Proportion 1 (p₁) or Percentage (%)", 
                                 value = "0.5",
                                 placeholder = "Enter as proportion (0.5) or percentage (50%)"),
                       div(
                         class = "who-info-box",
                         style = "margin-top: 5px; font-size: 12px;",
                         HTML("<strong>Flexible Input Format:</strong><br>
                              • Enter as proportion: <em>0.5</em> (for 50%)<br>
                              • Enter as percentage: <em>50%</em> or <em>50</em><br>
                              <em>Examples: 0.3, 30%, 5%, 0.05</em>")
                       ),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("p2_other", "Proportion 2 (p₂) or Percentage (%)", 
                                 value = "0.3",
                                 placeholder = "Enter as proportion (0.3) or percentage (30%)")
                     ),
                     "correlation" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("r_other", "Correlation Coefficient (r)", value = 0.3, min = -0.99, max = 0.99, step = 0.01)
                     ),
                     "regression" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("effect_size_other", "Effect Size (f²)", value = 0.15, min = 0.02, step = 0.01),
                       numericInput("predictors_other", "Number of Predictors", value = 1, min = 1, step = 1)
                     ),
                     "odds_ratio" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("or_other", "Odds Ratio (OR)", value = 2.0, min = 1.1, step = 0.1),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("p1_other", "Proportion in Control Group or Percentage (%)", 
                                 value = "0.2",
                                 placeholder = "Enter as proportion (0.2) or percentage (20%)")
                     ),
                     "relative_risk" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("rr_other", "Relative Risk (RR)", value = 1.5, min = 1.1, step = 0.1),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("p1_other", "Proportion in Control Group or Percentage (%)", 
                                 value = "0.2",
                                 placeholder = "Enter as proportion (0.2) or percentage (20%)")
                     ),
                     "prevalence" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("precision_other", "Precision (d)", value = 0.05, min = 0.01, step = 0.01),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("prevalence_other", "Expected Prevalence or Percentage (%)", 
                                 value = "0.1",
                                 placeholder = "Enter as proportion (0.1) or percentage (10%)")
                     ),
                     "case_control" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("or_other", "Odds Ratio (OR)", value = 2.0, min = 1.1, step = 0.1),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("p1_other", "Proportion in Controls or Percentage (%)", 
                                 value = "0.2",
                                 placeholder = "Enter as proportion (0.2) or percentage (20%)"),
                       numericInput("r_other", "Case:Control Ratio", value = 1, min = 0.1, step = 0.1)
                     ),
                     "cohort" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("rr_other", "Relative Risk (RR)", value = 1.5, min = 1.1, step = 0.1),
                       # UPDATED: Changed to textInput for percentage support
                       textInput("p1_other", "Proportion in Unexposed or Percentage (%)", 
                                 value = "0.2",
                                 placeholder = "Enter as proportion (0.2) or percentage (20%)"),
                       numericInput("r_other", "Exposed:Unexposed Ratio", value = 1, min = 0.1, step = 0.1)
                     )
    )
    
    return(params)
  })
  
  # UPDATED: Helper function to parse proportion inputs for Other Formulas
  parseProportionInput_other_formula <- function(input_str, param_name) {
    if (is.null(input_str) || input_str == "") {
      # Default values based on parameter
      if (param_name == "p_other") return(0.5)
      if (param_name == "p1_other") return(0.5)
      if (param_name == "p2_other") return(0.3)
      if (param_name == "prevalence_other") return(0.1)
      return(0.5)
    }
    
    return(parseProportionInput(input_str))
  }
  
  # UPDATED: Formula descriptions with Single Population Proportion
  output$formula_description <- renderUI({
    formula_type <- input$formula_type
    
    description <- switch(formula_type,
                          "single_proportion" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Single Population Proportion (EPI INFO 7.2.2.6):</strong><br>
                <em>n = Z² × p × (1-p) / d²</em><br>
                <ul>
                  <li><strong>Z</strong> = Z-score for desired confidence level</li>
                  <li><strong>p</strong> = estimated proportion (enter as proportion or percentage)</li>
                  <li><strong>d</strong> = margin of error</li>
                </ul>
                Use this formula for estimating a single proportion with specified precision.
                <p><strong>Note:</strong> You can enter proportion as 0.5 or as percentage 50% - both work the same!</p>
              ")
                          ),
                          "mean_known_var" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Mean Estimation with Known Variance:</strong><br>
                <em>n = (Z² × σ²) / d²</em><br>
                <ul>
                  <li><strong>Z</strong> = Z-score for desired confidence level</li>
                  <li><strong>σ</strong> = known standard deviation</li>
                  <li><strong>d</strong> = margin of error</li>
                </ul>
                Use this formula when the population standard deviation is known.
              ")
                          ),
                          "mean_unknown_var" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Mean Estimation with Unknown Variance:</strong><br>
                <em>n = (t² × s²) / d²</em> (approximated using power analysis)<br>
                <ul>
                  <li><strong>t</strong> = t-score for desired confidence level and degrees of freedom</li>
                  <li><strong>s</strong> = estimated standard deviation</li>
                  <li><strong>d</strong> = margin of error</li>
                </ul>
                Uses power analysis to determine sample size for t-test.
              ")
                          ),
                          "proportion_diff" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Difference Between Two Proportions:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (p₁ - p₂)²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>p₁, p₂</strong> = proportions in two groups (enter as proportion or percentage)</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For comparing two independent proportions.
                <p><strong>Note:</strong> You can enter proportions as 0.5 or as percentage 50% - both work the same!</p>
              ")
                          ),
                          "correlation" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Correlation Coefficient:</strong><br>
                <em>n = [(Zα + Zβ) / (0.5 × ln((1+r)/(1-r)))]² + 3</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>r</strong> = expected correlation coefficient</li>
                </ul>
                For testing significance of a correlation coefficient.
              ")
                          ),
                          "regression" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Regression Coefficient:</strong><br>
                <em>n = (Zα + Zβ)² / (f²) + k + 1</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>f²</strong> = effect size (R²/(1-R²))</li>
                  <li><strong>k</strong> = number of predictors</li>
                </ul>
                For testing significance of regression coefficients.
              ")
                          ),
                          "odds_ratio" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Odds Ratio:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (ln(OR))²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>OR</strong> = odds ratio</li>
                  <li><strong>p₁</strong> = proportion in control group (enter as proportion or percentage)</li>
                  <li><strong>p₂</strong> = p₁ × OR / (1 + p₁(OR-1))</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For case-control studies testing odds ratios.
                <p><strong>Note:</strong> You can enter proportion as 0.2 or as percentage 20% - both work the same!</p>
              ")
                          ),
                          "relative_risk" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Relative Risk:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (ln(RR))²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>RR</strong> = relative risk</li>
                  <li><strong>p₁</strong> = proportion in control group (enter as proportion or percentage)</li>
                  <li><strong>p₂</strong> = p₁ × RR</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For cohort studies testing relative risk.
                <p><strong>Note:</strong> You can enter proportion as 0.2 or as percentage 20% - both work the same!</p>
              ")
                          ),
                          "prevalence" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Prevalence Study:</strong><br>
                <em>n = (Z² × p × (1-p)) / d²</em><br>
                <ul>
                  <li><strong>Z</strong> = Z-score for desired confidence level</li>
                  <li><strong>p</strong> = expected prevalence (enter as proportion or percentage)</li>
                  <li><strong>d</strong> = precision (margin of error)</li>
                </ul>
                For estimating disease prevalence with specified precision.
                <p><strong>Note:</strong> You can enter prevalence as 0.1 or as percentage 10% - both work the same!</p>
              ")
                          ),
                          "case_control" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Case-Control Study:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (p₁ - p₂)²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>OR</strong> = odds ratio</li>
                  <li><strong>p₁</strong> = proportion in controls (enter as proportion or percentage)</li>
                  <li><strong>p₂</strong> = p₁ × OR / (1 + p₁(OR-1))</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For case-control studies with specified case:control ratio.
                <p><strong>Note:</strong> You can enter proportion as 0.2 or as percentage 20% - both work the same!</p>
              ")
                          ),
                          "cohort" = div(
                            class = "who-info-box",
                            HTML("
                <strong>Cohort Study:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (p₁ - p₂)²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>RR</strong> = relative risk</li>
                  <li><strong>p₁</strong> = proportion in unexposed (enter as proportion or percentage)</li>
                  <li><strong>p₂</strong> = p₁ × RR</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For cohort studies with specified exposed:unexposed ratio.
                <p><strong>Note:</strong> You can enter proportion as 0.2 or as percentage 20% - both work the same!</p>
              ")
                          )
    )
    
    return(description)
  })
  
  # UPDATED: Sample size calculation with Single Population Proportion and percentage support
  otherSampleSize <- reactive({
    formula_type <- input$formula_type
    
    tryCatch({
      result <- switch(formula_type,
                       "single_proportion" = {
                         alpha <- input$alpha_other
                         p <- parseProportionInput_other_formula(input$p_other, "p_other")
                         d <- input$d_other
                         z <- qnorm(1 - alpha/2)
                         ceiling((z^2 * p * (1 - p)) / (d^2))
                       },
                       "mean_known_var" = {
                         alpha <- input$alpha_other
                         sigma <- input$sigma_other
                         d <- input$d_other
                         z <- qnorm(1 - alpha/2)
                         ceiling((z^2 * sigma^2) / (d^2))
                       },
                       "mean_unknown_var" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         effect_size <- input$effect_size_other
                         result <- pwr.t.test(d = effect_size, sig.level = alpha, power = power, 
                                              type = "two.sample", alternative = "two.sided")
                         ceiling(result$n * 2)  # Total sample size for two groups
                       },
                       "proportion_diff" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                         p2 <- parseProportionInput_other_formula(input$p2_other, "p2_other")
                         result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                               alternative = "two.sided")
                         ceiling(result$n * 2)  # Total sample size for two groups
                       },
                       "correlation" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         r <- input$r_other
                         result <- pwr.r.test(r = r, sig.level = alpha, power = power, 
                                              alternative = "two.sided")
                         ceiling(result$n)
                       },
                       "regression" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         effect_size <- input$effect_size_other
                         predictors <- input$predictors_other
                         result <- pwr.f2.test(u = predictors, f2 = effect_size, 
                                               sig.level = alpha, power = power)
                         ceiling(result$v + predictors + 1)  # n = v + u + 1
                       },
                       "odds_ratio" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         or <- input$or_other
                         p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                         p2 <- p1 * or / (1 + p1 * (or - 1))
                         result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                               alternative = "two.sided")
                         ceiling(result$n * 2)  # Total sample size for two groups
                       },
                       "relative_risk" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         rr <- input$rr_other
                         p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                         p2 <- p1 * rr
                         result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                               alternative = "two.sided")
                         ceiling(result$n * 2)  # Total sample size for two groups
                       },
                       "prevalence" = {
                         alpha <- input$alpha_other
                         p <- parseProportionInput_other_formula(input$prevalence_other, "prevalence_other")
                         d <- input$precision_other
                         z <- qnorm(1 - alpha/2)
                         ceiling((z^2 * p * (1 - p)) / (d^2))
                       },
                       "case_control" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         or <- input$or_other
                         p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                         r <- input$r_other  # case:control ratio
                         
                         # FIXED: Ensure valid proportions and calculate p2 correctly
                         p1 <- max(0.01, min(0.99, p1))
                         p2 <- (p1 * or) / (1 + p1 * (or - 1))
                         p2 <- max(0.01, min(0.99, p2))
                         
                         # FIXED: Use proper sample size calculation for case-control
                         p_bar <- (p1 + p2) / 2
                         q_bar <- 1 - p_bar
                         q1 <- 1 - p1
                         q2 <- 1 - p2
                         
                         z_alpha <- qnorm(1 - alpha/2)
                         z_beta <- qnorm(power)
                         
                         n_per_group <- ((z_alpha * sqrt(2 * p_bar * q_bar) + 
                                            z_beta * sqrt(p1 * q1 + p2 * q2))^2) / ((p1 - p2)^2)
                         
                         n_cases <- ceiling(n_per_group)
                         n_controls <- ceiling(n_cases * r)
                         n_cases + n_controls  # Total sample size
                       },
                       "cohort" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         rr <- input$rr_other
                         p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                         r <- input$r_other  # exposed:unexposed ratio
                         
                         # FIXED: Ensure valid proportions and calculate p2 correctly
                         p1 <- max(0.01, min(0.99, p1))
                         p2 <- p1 * rr
                         p2 <- max(0.01, min(0.99, p2))
                         
                         # FIXED: Use proper sample size calculation for cohort
                         p_bar <- (p1 + p2) / 2
                         q_bar <- 1 - p_bar
                         q1 <- 1 - p1
                         q2 <- 1 - p2
                         
                         z_alpha <- qnorm(1 - alpha/2)
                         z_beta <- qnorm(power)
                         
                         n_per_group <- ((z_alpha * sqrt(2 * p_bar * q_bar) + 
                                            z_beta * sqrt(p1 * q1 + p2 * q2))^2) / ((p1 - p2)^2)
                         
                         n_exposed <- ceiling(n_per_group)
                         n_unexposed <- ceiling(n_exposed * r)
                         n_exposed + n_unexposed  # Total sample size
                       }
      )
      
      if (is.null(result) || is.na(result)) return(0)
      result
    }, error = function(e) {
      message("Error in sample size calculation: ", e$message)
      return(0)
    })
  })
  
  adjustedSampleSize_other <- reactive({
    n <- otherSampleSize()
    if (n == 0) return(0)
    non_response_rate <- input$non_response_other / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  # UPDATED: Formula explanations with Single Population Proportion and percentage support
  output$formulaExplanation_other <- renderText({
    formula_type <- input$formula_type
    
    explanation <- switch(formula_type,
                          "single_proportion" = {
                            alpha <- input$alpha_other
                            p <- parseProportionInput_other_formula(input$p_other, "p_other")
                            d <- input$d_other
                            z <- qnorm(1 - alpha/2)
                            n <- ceiling((z^2 * p * (1 - p)) / (d^2))
                            paste(
                              "Single Population Proportion Formula (EPI INFO 7.2.2.6):\n",
                              paste("Input provided: '", input$p_other, "'", sep = ""),
                              paste("Interpreted as proportion: ", round(p, 4), " (", round(p*100, 2), "%)", sep = ""),
                              "\nFormula: n = Z² × p × (1-p) / d²\n",
                              "Calculation: (", round(z, 3), "² × ", round(p, 4), " × (1-", round(p, 4), ")) / ", d, "²\n",
                              "Step-by-step: (", round(z^2, 3), " × ", round(p, 4), " × ", round((1-p), 4), ") / ", d^2, "\n",
                              "Result: ", (z^2 * p * (1-p)), " / ", d^2, " = ", (z^2 * p * (1-p)) / d^2, "\n",
                              "Apply ceiling function: ", n
                            )
                          },
                          "mean_known_var" = {
                            alpha <- input$alpha_other
                            sigma <- input$sigma_other
                            d <- input$d_other
                            z <- qnorm(1 - alpha/2)
                            paste(
                              "Mean Estimation with Known Variance:\n",
                              "Formula: n = (Z² × σ²) / d²\n",
                              "Calculation: (", round(z, 3), "² × ", sigma, "²) / ", d, "²\n",
                              "Step-by-step: (", round(z^2, 3), " × ", sigma^2, ") / ", d^2, "\n",
                              "Result: ", (z^2 * sigma^2), " / ", d^2, " = ", (z^2 * sigma^2) / d^2, "\n",
                              "Apply ceiling function: ", otherSampleSize()
                            )
                          },
                          "mean_unknown_var" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            effect_size <- input$effect_size_other
                            result <- pwr.t.test(d = effect_size, sig.level = alpha, power = power, 
                                                 type = "two.sample", alternative = "two.sided")
                            paste(
                              "Mean Estimation with Unknown Variance:\n",
                              "Using power analysis for two-sample t-test\n",
                              "Effect size (Cohen's d): ", effect_size, "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Sample size per group: ", ceiling(result$n), "\n",
                              "Total sample size: ", ceiling(result$n * 2)
                            )
                          },
                          "proportion_diff" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                            p2 <- parseProportionInput_other_formula(input$p2_other, "p2_other")
                            result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                                  alternative = "two.sided")
                            paste(
                              "Difference Between Two Proportions:\n",
                              paste("Proportion 1 input: '", input$p1_other, "'", sep = ""),
                              paste("Interpreted as: ", round(p1, 4), " (", round(p1*100, 2), "%)", sep = ""),
                              paste("Proportion 2 input: '", input$p2_other, "'", sep = ""),
                              paste("Interpreted as: ", round(p2, 4), " (", round(p2*100, 2), "%)", sep = ""),
                              "Effect size (h): ", round(ES.h(p1, p2), 3), "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Sample size per group: ", ceiling(result$n), "\n",
                              "Total sample size: ", ceiling(result$n * 2)
                            )
                          },
                          "correlation" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            r <- input$r_other
                            result <- pwr.r.test(r = r, sig.level = alpha, power = power, 
                                                 alternative = "two.sided")
                            paste(
                              "Correlation Coefficient:\n",
                              "Correlation (r): ", r, "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Required sample size: ", ceiling(result$n)
                            )
                          },
                          "regression" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            effect_size <- input$effect_size_other
                            predictors <- input$predictors_other
                            result <- pwr.f2.test(u = predictors, f2 = effect_size, 
                                                  sig.level = alpha, power = power)
                            paste(
                              "Regression Coefficient:\n",
                              "Effect size (f²): ", effect_size, "\n",
                              "Number of predictors: ", predictors, "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Required sample size: n = v + u + 1 = ", 
                              ceiling(result$v), " + ", predictors, " + 1 = ", 
                              ceiling(result$v + predictors + 1)
                            )
                          },
                          "odds_ratio" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            or <- input$or_other
                            p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                            p2 <- p1 * or / (1 + p1 * (or - 1))
                            result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                                  alternative = "two.sided")
                            paste(
                              "Odds Ratio:\n",
                              paste("Control proportion input: '", input$p1_other, "'", sep = ""),
                              paste("Interpreted as: ", round(p1, 4), " (", round(p1*100, 2), "%)", sep = ""),
                              "Odds ratio: ", or, "\n",
                              "Case proportion: ", round(p2, 3), " (", round(p2*100, 2), "%)", "\n",
                              "Effect size (h): ", round(ES.h(p1, p2), 3), "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Sample size per group: ", ceiling(result$n), "\n",
                              "Total sample size: ", ceiling(result$n * 2)
                            )
                          },
                          "relative_risk" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            rr <- input$rr_other
                            p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                            p2 <- p1 * rr
                            result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                                  alternative = "two.sided")
                            paste(
                              "Relative Risk:\n",
                              paste("Unexposed proportion input: '", input$p1_other, "'", sep = ""),
                              paste("Interpreted as: ", round(p1, 4), " (", round(p1*100, 2), "%)", sep = ""),
                              "Relative risk: ", rr, "\n",
                              "Exposed proportion: ", round(p2, 3), " (", round(p2*100, 2), "%)", "\n",
                              "Effect size (h): ", round(ES.h(p1, p2), 3), "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Sample size per group: ", ceiling(result$n), "\n",
                              "Total sample size: ", ceiling(result$n * 2)
                            )
                          },
                          "prevalence" = {
                            alpha <- input$alpha_other
                            p <- parseProportionInput_other_formula(input$prevalence_other, "prevalence_other")
                            d <- input$precision_other
                            z <- qnorm(1 - alpha/2)
                            paste(
                              "Prevalence Study:\n",
                              paste("Prevalence input: '", input$prevalence_other, "'", sep = ""),
                              paste("Interpreted as: ", round(p, 4), " (", round(p*100, 2), "%)", sep = ""),
                              "\nFormula: n = (Z² × p × (1-p)) / d²\n",
                              "Calculation: (", round(z, 3), "² × ", round(p, 4), " × (1-", round(p, 4), ")) / ", d, "²\n",
                              "Step-by-step: (", round(z^2, 3), " × ", round(p, 4), " × ", round((1-p), 4), ") / ", d^2, "\n",
                              "Result: ", (z^2 * p * (1-p)), " / ", d^2, " = ", (z^2 * p * (1-p)) / d^2, "\n",
                              "Apply ceiling function: ", otherSampleSize()
                            )
                          },
                          "case_control" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            or <- input$or_other
                            p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                            r <- input$r_other
                            
                            # FIXED calculations
                            p1 <- max(0.01, min(0.99, p1))
                            p2 <- (p1 * or) / (1 + p1 * (or - 1))
                            p2 <- max(0.01, min(0.99, p2))
                            
                            p_bar <- (p1 + p2) / 2
                            q_bar <- 1 - p_bar
                            q1 <- 1 - p1
                            q2 <- 1 - p2
                            
                            z_alpha <- qnorm(1 - alpha/2)
                            z_beta <- qnorm(power)
                            
                            n_per_group <- ((z_alpha * sqrt(2 * p_bar * q_bar) + 
                                               z_beta * sqrt(p1 * q1 + p2 * q2))^2) / ((p1 - p2)^2)
                            
                            n_cases <- ceiling(n_per_group)
                            n_controls <- ceiling(n_cases * r)
                            total_n <- n_cases + n_controls
                            
                            paste(
                              "Case-Control Study (Fixed):\n",
                              paste("Control proportion input: '", input$p1_other, "'", sep = ""),
                              paste("Interpreted as: ", round(p1, 4), " (", round(p1*100, 2), "%)", sep = ""),
                              "Odds ratio: ", or, "\n",
                              "Case proportion: ", round(p2, 3), " (", round(p2*100, 2), "%)", "\n",
                              "Case:Control ratio: ", r, "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Cases needed: ", n_cases, "\n",
                              "Controls needed: ", n_controls, "\n",
                              "Total sample size: ", total_n
                            )
                          },
                          "cohort" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            rr <- input$rr_other
                            p1 <- parseProportionInput_other_formula(input$p1_other, "p1_other")
                            r <- input$r_other
                            
                            # FIXED calculations
                            p1 <- max(0.01, min(0.99, p1))
                            p2 <- p1 * rr
                            p2 <- max(0.01, min(0.99, p2))
                            
                            p_bar <- (p1 + p2) / 2
                            q_bar <- 1 - p_bar
                            q1 <- 1 - p1
                            q2 <- 1 - p2
                            
                            z_alpha <- qnorm(1 - alpha/2)
                            z_beta <- qnorm(power)
                            
                            n_per_group <- ((z_alpha * sqrt(2 * p_bar * q_bar) + 
                                               z_beta * sqrt(p1 * q1 + p2 * q2))^2) / ((p1 - p2)^2)
                            
                            n_exposed <- ceiling(n_per_group)
                            n_unexposed <- ceiling(n_exposed * r)
                            total_n <- n_exposed + n_unexposed
                            
                            paste(
                              "Cohort Study (Fixed):\n",
                              paste("Unexposed proportion input: '", input$p1_other, "'", sep = ""),
                              paste("Interpreted as: ", round(p1, 4), " (", round(p1*100, 2), "%)", sep = ""),
                              "Relative risk: ", rr, "\n",
                              "Exposed proportion: ", round(p2, 3), " (", round(p2*100, 2), "%)", "\n",
                              "Exposed:Unexposed ratio: ", r, "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Exposed needed: ", n_exposed, "\n",
                              "Unexposed needed: ", n_unexposed, "\n",
                              "Total sample size: ", total_n
                            )
                          }
    )
    
    return(explanation)
  })
  
  output$sampleSize_other <- renderText({
    paste("Required Sample Size:", otherSampleSize())
  })
  
  output$adjustedSampleSize_other <- renderText({
    n <- otherSampleSize()
    adj_n <- adjustedSampleSize_other()
    non_response_rate <- input$non_response_other
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", n, "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", n, " / (1 - ", non_response_rate/100, ") = ", n / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  allocationData_other <- reactive({
    N <- totalPopulation_other()
    n <- adjustedSampleSize_other()
    strata <- sapply(1:rv_other$stratumCount, function(i) input[[paste0("stratum_other", i)]] )
    pops <- sapply(1:rv_other$stratumCount, function(i) input[[paste0("pop_other", i)]] )
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv_other$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  interpretationText_other <- reactive({
    alloc <- allocationData_other()
    if (is.null(alloc)) return("")
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    paste0("Based on the ", input$formula_type, " formula",
           " and a non-response rate of ", input$non_response_other, "%",
           ", the required sample size is ", adjustedSampleSize_other(), ". ",
           paste(sentences, collapse = " "))
  })
  
  output$allocationTable_other <- renderTable({ allocationData_other() })
  output$interpretationText_other <- renderText({ interpretationText_other() })
  
  # UPDATED: Power Analysis with robust error handling for Simple Linear Regression
  observeEvent(input$runPower, {
    tryCatch({
      # Input validation
      if (is.na(input$effectSize) || input$effectSize <= 0) {
        stop("Effect size must be a positive number")
      }
      
      if (is.na(input$alpha) || input$alpha <= 0 || input$alpha >= 1) {
        stop("Alpha must be between 0 and 1")
      }
      
      if (is.na(input$power) || input$power <= 0 || input$power >= 1) {
        stop("Power must be between 0 and 1")
      }
      
      result <- switch(input$testType,
                       "Independent t-test" = {
                         pwr.t.test(d = input$effectSize, sig.level = input$alpha, 
                                    power = input$power, type = "two.sample", 
                                    alternative = "two.sided")
                       },
                       "Paired t-test" = {
                         pwr.t.test(d = input$effectSize, sig.level = input$alpha, 
                                    power = input$power, type = "paired", 
                                    alternative = "two.sided")
                       },
                       "One-sample t-test" = {
                         pwr.t.test(d = input$effectSize, sig.level = input$alpha, 
                                    power = input$power, type = "one.sample", 
                                    alternative = "two.sided")
                       },
                       "One-Way ANOVA" = {
                         k <- ifelse(!is.null(input$groups) && !is.na(input$groups), 
                                     input$groups, 3)
                         pwr.anova.test(k = k, f = input$effectSize, sig.level = input$alpha, 
                                        power = input$power)
                       },
                       "Two-Way ANOVA" = {
                         k <- ifelse(!is.null(input$groups) && !is.na(input$groups), 
                                     input$groups, 4)
                         pwr.anova.test(k = k, f = input$effectSize, sig.level = input$alpha, 
                                        power = input$power)
                       },
                       "Proportion" = {
                         p1 <- input$effectSize
                         p2 <- 0.5
                         
                         if (p1 <= 0 || p1 >= 1) {
                           stop("Proportion must be between 0 and 1")
                         }
                         
                         h_val <- ES.h(p1, p2)
                         if (is.na(h_val) || !is.finite(h_val)) {
                           stop("Invalid effect size for proportion test")
                         }
                         pwr.2p.test(h = h_val, sig.level = input$alpha, 
                                     power = input$power, alternative = "two.sided")
                       },
                       "Correlation" = {
                         r_val <- input$effectSize
                         if (abs(r_val) >= 1) {
                           stop("Correlation coefficient must be between -1 and 1")
                         }
                         pwr.r.test(r = r_val, sig.level = input$alpha, 
                                    power = input$power, alternative = "two.sided")
                       },
                       "Chi-squared" = {
                         w_val <- input$effectSize
                         df_val <- ifelse(!is.null(input$df) && !is.na(input$df), 
                                          input$df, 1)
                         
                         if (w_val <= 0 || w_val > 1) {
                           stop("Effect size for chi-squared (w) must be between 0 and 1")
                         }
                         if (df_val <= 0) {
                           stop("Degrees of freedom must be positive")
                         }
                         
                         pwr.chisq.test(w = w_val, df = df_val, sig.level = input$alpha, 
                                        power = input$power)
                       },
                       "Simple Linear Regression" = {
                         # ROBUST FIX: Enhanced validation for Simple Linear Regression
                         f2_val <- input$effectSize
                         
                         if (f2_val <= 0) {
                           stop("Effect size (f²) must be positive for regression")
                         }
                         
                         if (f2_val > 10) {
                           warning("Very large effect size detected. This may indicate unrealistic expectations.")
                         }
                         
                         # For simple linear regression: u = 1 (one predictor)
                         result <- pwr.f2.test(u = 1, f2 = f2_val, 
                                               sig.level = input$alpha, power = input$power)
                         
                         # Calculate total sample size: n = v + u + 1
                         # Where v = denominator degrees of freedom from pwr.f2.test
                         # u = number of predictors (1 for simple regression)
                         # +1 accounts for the intercept
                         total_n <- ceiling(result$v + 1 + 1)
                         
                         # Return both the pwr result and calculated total sample size
                         list(pwr_result = result, total_n = total_n)
                       },
                       "Multiple Linear Regression" = {
                         f2_val <- input$effectSize
                         predictors <- ifelse(!is.null(input$predictors) && !is.na(input$predictors), 
                                              input$predictors, 2)
                         
                         if (f2_val <= 0) {
                           stop("Effect size (f²) must be positive for regression")
                         }
                         
                         if (predictors <= 0) {
                           stop("Number of predictors must be positive")
                         }
                         
                         result <- pwr.f2.test(u = predictors, f2 = f2_val, 
                                               sig.level = input$alpha, power = input$power)
                         
                         # Calculate total sample size: n = v + u + 1
                         total_n <- ceiling(result$v + predictors + 1)
                         
                         list(pwr_result = result, total_n = total_n)
                       }
      )
      
      # Clear any previous errors
      output$powerError <- renderText({ "" })
      outputOptions(output, "powerError", suspendWhenHidden = FALSE)
      
      # Format results based on test type
      if (input$testType %in% c("Simple Linear Regression", "Multiple Linear Regression")) {
        sample_size <- result$total_n
        pwr_result <- result$pwr_result
      } else {
        sample_size <- if(input$testType %in% c("Independent t-test", "Paired t-test", "Proportion")) {
          ceiling(result$n * 2)
        } else if (input$testType %in% c("One-Way ANOVA", "Two-Way ANOVA")) {
          ceiling(result$n)
        } else if (input$testType == "Chi-squared") {
          ceiling(result$N)
        } else {
          ceiling(result$n)
        }
        pwr_result <- result
      }
      
      per_group <- if(input$testType %in% c("Independent t-test", "Paired t-test", "Proportion")) {
        ceiling(pwr_result$n)
      } else if (input$testType %in% c("One-Way ANOVA", "Two-Way ANOVA")) {
        ceiling(pwr_result$n)
      } else {
        NULL
      }
      
      output$powerResult <- renderText({
        result_text <- paste(
          "Power Analysis Results for", input$testType, "\n",
          "========================================\n",
          "Effect size:", round(input$effectSize, 3), 
          if(input$testType %in% c("Simple Linear Regression", "Multiple Linear Regression")) 
            paste("(f² =", round(input$effectSize, 3), ")") 
          else if(input$testType == "Chi-squared") paste("(w =", round(input$effectSize, 3), ")") 
          else if(input$testType == "Correlation") paste("(r =", round(input$effectSize, 3), ")") 
          else "",
          "\n",
          "Alpha:", input$alpha, "\n",
          "Power:", input$power, "\n",
          "Required sample size:", sample_size, "\n"
        )
        
        if (!is.null(per_group)) {
          result_text <- paste(
            result_text,
            if(input$testType %in% c("Independent t-test", "Paired t-test", "Proportion")) {
              paste("Sample size per group:", per_group)
            } else if(input$testType %in% c("One-Way ANOVA", "Two-Way ANOVA")) {
              paste("Sample size per group:", per_group)
            } else ""
          )
        }
        
        if(input$testType == "Multiple Linear Regression") {
          result_text <- paste(result_text, "\nNumber of predictors:", input$predictors)
        }
        
        if(input$testType == "Simple Linear Regression") {
          result_text <- paste(result_text, "\nPredictors: 1 (simple linear regression)")
          
          # Add R² interpretation for better understanding
          r_squared <- input$effectSize / (1 + input$effectSize)
          result_text <- paste(result_text, 
                               "\nEquivalent R²:", round(r_squared, 4),
                               "(" , round(r_squared * 100, 1), "% of variance explained)")
        }
        
        if(input$testType == "Chi-squared") {
          result_text <- paste(result_text, "\nDegrees of freedom:", input$df)
        }
        
        result_text
      })
      
    }, error = function(e) {
      # Display user-friendly error message
      output$powerResult <- renderText({ 
        paste("Power analysis could not be completed.\n",
              "Please check your input values and try again.")
      })
      
      output$powerError <- renderText({
        paste("Error:", e$message)
      })
      outputOptions(output, "powerError", suspendWhenHidden = FALSE)
    })
  })
  
  # Enhanced Power Analysis download handler with Simple Linear Regression support
  output$downloadPowerSteps <- downloadHandler(
    filename = function() {
      paste("power_analysis_results_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      tryCatch({
        # Replicate the power calculation for the download
        result <- switch(input$testType,
                         "Independent t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, 
                                                           power = input$power, type = "two.sample", 
                                                           alternative = "two.sided"),
                         "Paired t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, 
                                                      power = input$power, type = "paired", 
                                                      alternative = "two.sided"),
                         "One-sample t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, 
                                                          power = input$power, type = "one.sample", 
                                                          alternative = "two.sided"),
                         "One-Way ANOVA" = {
                           k <- ifelse(!is.null(input$groups) && !is.na(input$groups), 
                                       input$groups, 3)
                           pwr.anova.test(k = k, f = input$effectSize, sig.level = input$alpha, 
                                          power = input$power)
                         },
                         "Two-Way ANOVA" = {
                           k <- ifelse(!is.null(input$groups) && !is.na(input$groups), 
                                       input$groups, 4)
                           pwr.anova.test(k = k, f = input$effectSize, sig.level = input$alpha, 
                                          power = input$power)
                         },
                         "Proportion" = {
                           p1 <- input$effectSize
                           p2 <- 0.5
                           h_val <- ES.h(p1, p2)
                           pwr.2p.test(h = h_val, sig.level = input$alpha, 
                                       power = input$power, alternative = "two.sided")
                         },
                         "Correlation" = pwr.r.test(r = input$effectSize, sig.level = input$alpha, 
                                                    power = input$power, alternative = "two.sided"),
                         "Chi-squared" = {
                           w_val <- input$effectSize
                           df_val <- ifelse(!is.null(input$df) && !is.na(input$df), 
                                            input$df, 1)
                           pwr.chisq.test(w = w_val, df = df_val, sig.level = input$alpha, 
                                          power = input$power)
                         },
                         "Simple Linear Regression" = {
                           f2_val <- input$effectSize
                           result <- pwr.f2.test(u = 1, f2 = f2_val, 
                                                 sig.level = input$alpha, power = input$power)
                           total_n <- ceiling(result$v + 1 + 1)
                           list(pwr_result = result, total_n = total_n)
                         },
                         "Multiple Linear Regression" = {
                           f2_val <- input$effectSize
                           predictors <- ifelse(!is.null(input$predictors) && !is.na(input$predictors), 
                                                input$predictors, 2)
                           result <- pwr.f2.test(u = predictors, f2 = f2_val, 
                                                 sig.level = input$alpha, power = input$power)
                           total_n <- ceiling(result$v + predictors + 1)
                           list(pwr_result = result, total_n = total_n)
                         }
        )
        
        # Calculate sample size based on test type
        if (input$testType %in% c("Simple Linear Regression", "Multiple Linear Regression")) {
          sample_size <- result$total_n
          pwr_result <- result$pwr_result
        } else {
          sample_size <- if(input$testType %in% c("Independent t-test", "Paired t-test", "Proportion")) {
            ceiling(result$n * 2)
          } else if (input$testType %in% c("One-Way ANOVA", "Two-Way ANOVA")) {
            ceiling(result$n)
          } else if (input$testType == "Chi-squared") {
            ceiling(result$N)
          } else {
            ceiling(result$n)
          }
          pwr_result <- result
        }
        
        steps <- c(
          "POWER ANALYSIS RESULTS",
          "======================",
          paste("Date:", Sys.Date()),
          paste("Statistical Test:", input$testType),
          paste("Effect Size:", round(input$effectSize, 4)),
          paste("Alpha (Significance Level):", input$alpha),
          paste("Desired Power:", input$power),
          ""
        )
        
        # Add test-specific parameters
        if(input$testType == "Multiple Linear Regression") {
          steps <- c(steps, paste("Number of Predictors:", input$predictors))
        }
        
        if(input$testType == "Simple Linear Regression") {
          steps <- c(steps, "Number of Predictors: 1")
          
          # Add R² conversion for interpretation
          r_squared <- input$effectSize / (1 + input$effectSize)
          steps <- c(steps, 
                     paste("Equivalent R²:", round(r_squared, 4)),
                     paste("Variance Explained:", round(r_squared * 100, 1), "%"))
        }
        
        if(input$testType == "Chi-squared") {
          steps <- c(steps, paste("Degrees of Freedom:", input$df))
        }
        
        if(input$testType %in% c("One-Way ANOVA", "Two-Way ANOVA")) {
          steps <- c(steps, paste("Number of Groups:", input$groups))
        }
        
        steps <- c(steps,
                   "",
                   "RESULTS:",
                   paste("Required sample size:", sample_size)
        )
        
        if(input$testType %in% c("Independent t-test", "Paired t-test", "Proportion")) {
          steps <- c(steps, paste("Sample size per group:", ceiling(pwr_result$n)))
        }
        
        if(input$testType %in% c("One-Way ANOVA", "Two-Way ANOVA")) {
          steps <- c(steps, paste("Sample size per group:", ceiling(pwr_result$n)))
        }
        
        # Add formula explanation for Simple Linear Regression
        if(input$testType == "Simple Linear Regression") {
          steps <- c(steps,
                     "",
                     "CALCULATION DETAILS:",
                     "Formula: n = v + u + 1",
                     "Where:",
                     "- v = denominator degrees of freedom from power calculation",
                     "- u = number of predictors (1 for simple linear regression)",
                     "- +1 accounts for the intercept parameter",
                     paste("Calculation:", ceiling(pwr_result$v), "+ 1 + 1 =", sample_size)
          )
        }
        
        steps <- c(steps,
                   "",
                   "INTERPRETATION:",
                   paste("To achieve", input$power * 100, "% power to detect an effect size of", 
                         round(input$effectSize, 3), "with a significance level of", input$alpha * 100, "%,"),
                   paste("you need a total sample size of", sample_size, "participants."),
                   "",
                   "Note: This calculation assumes the specified parameters and may need adjustment",
                   "based on your specific research design and assumptions."
        )
        
        writeLines(steps, file)
        
      }, error = function(e) {
        error_steps <- c(
          "POWER ANALYSIS RESULTS",
          "======================",
          paste("Date:", Sys.Date()),
          paste("Error:", e$message),
          "",
          "Please check your input values and try again.",
          "",
          "For Simple Linear Regression:",
          "- Effect size (f²) must be positive",
          "- f² = R² / (1 - R²)",
          "- Common values: small=0.02, medium=0.15, large=0.35"
        )
        writeLines(error_steps, file)
      })
    }
  )
  
  # Descriptive Statistics Section
  output$dataTypeDetection <- renderUI({
    if (input$dataInput == "") {
      return(p("Please enter data to analyze."))
    }
    
    data_text <- input$dataInput
    lines <- strsplit(data_text, "\n")[[1]]
    
    if (length(lines) < 2) {
      return(p("Please enter at least one data value."))
    }
    
    # Try to detect if there's a header
    first_line <- lines[1]
    second_line <- lines[2]
    
    # Check if first line might be a header (contains non-numeric characters)
    has_header <- grepl("[A-Za-z]", first_line) && !grepl("^\\s*\\d+\\s*$", first_line)
    
    if (has_header) {
      data_values <- lines[-1]
    } else {
      data_values <- lines
    }
    
    # Clean and parse data
    clean_data <- gsub("\\s+", " ", data_values)
    clean_data <- trimws(clean_data)
    clean_data <- clean_data[clean_data != ""]
    
    if (length(clean_data) == 0) {
      return(p("No valid data found."))
    }
    
    # Try to parse as numeric first
    numeric_data <- suppressWarnings(as.numeric(clean_data))
    num_numeric <- sum(!is.na(numeric_data))
    num_total <- length(clean_data)
    
    if (num_numeric == num_total) {
      data_type <- "numerical"
      detection_text <- paste("Detected: Numerical/Continuous data (", num_total, " values)", sep = "")
    } else if (num_numeric == 0) {
      data_type <- "categorical"
      detection_text <- paste("Detected: Categorical data (", num_total, " values)", sep = "")
    } else {
      data_type <- "mixed"
      detection_text <- paste("Detected: Mixed data (", num_numeric, " numerical, ", 
                              num_total - num_numeric, " categorical)", sep = "")
    }
    
    # Override with user selection if not auto
    if (input$dataType != "auto") {
      data_type <- input$dataType
      detection_text <- paste("User specified:", 
                              switch(input$dataType,
                                     "numerical" = "Numerical/Continuous",
                                     "nominal" = "Categorical (Nominal)",
                                     "ordinal" = "Ordinal"))
    }
    
    output$dataType <- reactive({ data_type })
    outputOptions(output, "dataType", suspendWhenHidden = FALSE)
    
    tagList(
      p(strong(detection_text)),
      p("Analysis will provide appropriate statistics for this data type.")
    )
  })
  
  # Reactive for parsed data
  parsedData <- reactive({
    if (input$dataInput == "") return(NULL)
    
    data_text <- input$dataInput
    lines <- strsplit(data_text, "\n")[[1]]
    
    # Detect header
    first_line <- lines[1]
    has_header <- grepl("[A-Za-z]", first_line) && !grepl("^\\s*\\d+\\s*$", first_line)
    
    if (has_header) {
      data_values <- lines[-1]
    } else {
      data_values <- lines
    }
    
    # Clean data
    clean_data <- gsub("\\s+", " ", data_values)
    clean_data <- trimws(clean_data)
    clean_data <- clean_data[clean_data != ""]
    
    # Parse based on detected type
    data_type <- input$dataType
    if (data_type == "auto") {
      numeric_data <- suppressWarnings(as.numeric(clean_data))
      num_numeric <- sum(!is.na(numeric_data))
      if (num_numeric == length(clean_data)) {
        data_type <- "numerical"
      } else if (num_numeric == 0) {
        data_type <- "categorical"
      } else {
        data_type <- "mixed"
      }
    }
    
    list(
      values = clean_data,
      type = data_type,
      has_header = has_header,
      header = if(has_header) first_line else NULL
    )
  })
  
  # Numerical data summary
  output$numericalSummary <- renderText({
    data <- parsedData()
    if (is.null(data) || data$type != "numerical") return("No numerical data to analyze.")
    
    values <- as.numeric(data$values)
    values <- values[!is.na(values)]
    
    if (length(values) == 0) return("No valid numerical values found.")
    
    n <- length(values)
    mean_val <- mean(values)
    median_val <- median(values)
    sd_val <- sd(values)
    var_val <- var(values)
    min_val <- min(values)
    max_val <- max(values)
    range_val <- max_val - min_val
    q1 <- quantile(values, 0.25)
    q3 <- quantile(values, 0.75)
    iqr <- q3 - q1
    
    # Skewness and Kurtosis
    skewness_val <- e1071::skewness(values, type = 2)
    kurtosis_val <- e1071::kurtosis(values, type = 2)
    
    paste(
      "NUMERICAL DATA SUMMARY",
      "======================",
      paste("Sample size (n):", n),
      paste("Mean:", round(mean_val, 4)),
      paste("Median:", round(median_val, 4)),
      paste("Standard Deviation:", round(sd_val, 4)),
      paste("Variance:", round(var_val, 4)),
      paste("Minimum:", round(min_val, 4)),
      paste("Maximum:", round(max_val, 4)),
      paste("Range:", round(range_val, 4)),
      paste("1st Quartile (Q1):", round(q1, 4)),
      paste("3rd Quartile (Q3):", round(q3, 4)),
      paste("Interquartile Range (IQR):", round(iqr, 4)),
      paste("Skewness:", round(skewness_val, 4)),
      paste("Kurtosis:", round(kurtosis_val, 4)),
      "",
      "INTERPRETATION:",
      if(skewness_val > 0.5) "Data is positively skewed (right-skewed)" 
      else if(skewness_val < -0.5) "Data is negatively skewed (left-skewed)"
      else "Data is approximately symmetric",
      if(kurtosis_val > 0) "Distribution is leptokurtic (heavy-tailed)"
      else if(kurtosis_val < 0) "Distribution is platykurtic (light-tailed)"
      else "Distribution is mesokurtic (normal tails)",
      sep = "\n"
    )
  })
  
  # Categorical data summary
  output$categoricalSummary <- renderText({
    data <- parsedData()
    if (is.null(data) || !data$type %in% c("categorical", "nominal")) return("No categorical data to analyze.")
    
    values <- data$values
    freq_table <- table(values)
    prop_table <- prop.table(freq_table)
    
    n <- length(values)
    n_categories <- length(freq_table)
    mode_val <- names(freq_table)[which.max(freq_table)]
    mode_freq <- max(freq_table)
    
    summary_text <- paste(
      "CATEGORICAL DATA SUMMARY",
      "========================",
      paste("Sample size (n):", n),
      paste("Number of categories:", n_categories),
      paste("Mode:", mode_val, "(frequency:", mode_freq, ")"),
      "",
      "FREQUENCY DISTRIBUTION:",
      sep = "\n"
    )
    
    # Add frequency table
    for (i in 1:length(freq_table)) {
      category <- names(freq_table)[i]
      freq <- freq_table[i]
      prop <- round(prop_table[i] * 100, 2)
      summary_text <- paste(summary_text, 
                            paste(category, ": ", freq, " (", prop, "%)", sep = ""), 
                            sep = "\n")
    }
    
    summary_text
  })
  
  # Ordinal data summary
  output$ordinalSummary <- renderText({
    data <- parsedData()
    if (is.null(data) || data$type != "ordinal") return("No ordinal data to analyze.")
    
    values <- data$values
    freq_table <- table(values)
    prop_table <- prop.table(freq_table)
    
    n <- length(values)
    n_categories <- length(freq_table)
    mode_val <- names(freq_table)[which.max(freq_table)]
    
    # For ordinal data, we can calculate median-like statistics
    ordered_levels <- unique(values)
    
    summary_text <- paste(
      "ORDINAL DATA SUMMARY",
      "====================",
      paste("Sample size (n):", n),
      paste("Number of categories:", n_categories),
      paste("Mode:", mode_val),
      "",
      "FREQUENCY DISTRIBUTION:",
      sep = "\n"
    )
    
    # Add frequency table
    for (i in 1:length(freq_table)) {
      category <- names(freq_table)[i]
      freq <- freq_table[i]
      prop <- round(prop_table[i] * 100, 2)
      summary_text <- paste(summary_text, 
                            paste(category, ": ", freq, " (", prop, "%)", sep = ""), 
                            sep = "\n")
    }
    
    summary_text
  })
  
  # Mixed data summary
  output$mixedSummary <- renderText({
    data <- parsedData()
    if (is.null(data) || data$type != "mixed") return("No mixed data to analyze.")
    
    values <- data$values
    numeric_vals <- suppressWarnings(as.numeric(values))
    numeric_vals <- numeric_vals[!is.na(numeric_vals)]
    categorical_vals <- values[is.na(suppressWarnings(as.numeric(values)))]
    
    n_total <- length(values)
    n_numeric <- length(numeric_vals)
    n_categorical <- length(categorical_vals)
    
    paste(
      "MIXED DATA SUMMARY",
      "==================",
      paste("Total observations:", n_total),
      paste("Numerical values:", n_numeric, "(", round(n_numeric/n_total*100, 1), "%)"),
      paste("Categorical values:", n_categorical, "(", round(n_categorical/n_total*100, 1), "%)"),
      "",
      "Please separate numerical and categorical data for proper analysis.",
      "You can use the data type selector to force a specific analysis.",
      sep = "\n"
    )
  })
  
  # Unknown data type
  output$unknownData <- renderText({
    "Please enter data to analyze or select a specific data type."
  })
  
  # Plots for numerical data
  output$numericalPlots <- renderPlot({
    data <- parsedData()
    if (is.null(data) || data$type != "numerical") return(NULL)
    
    values <- as.numeric(data$values)
    values <- values[!is.na(values)]
    
    if (length(values) < 2) return(NULL)
    
    par(mfrow = c(1, 2))
    
    # Histogram
    hist(values, main = "Histogram", xlab = "Values", 
         col = "#6BAED6", border = "white", freq = FALSE)
    lines(density(values), col = "#00689D", lwd = 2)
    
    # Boxplot
    boxplot(values, main = "Boxplot", col = "#6BAED6", 
            ylab = "Values", horizontal = TRUE)
    
    par(mfrow = c(1, 1))
  })
  
  # Plots for categorical data
  output$categoricalPlots <- renderPlot({
    data <- parsedData()
    if (is.null(data) || !data$type %in% c("categorical", "nominal")) return(NULL)
    
    values <- data$values
    freq_table <- table(values)
    
    if (length(freq_table) == 0) return(NULL)
    
    par(mfrow = c(1, 2))
    
    # Bar plot
    barplot(freq_table, main = "Bar Plot", col = "#6BAED6", 
            las = 2, cex.names = 0.8)
    
    # Pie chart
    pie(freq_table, main = "Pie Chart", col = rainbow(length(freq_table)))
    
    par(mfrow = c(1, 1))
  })
  
  # Plots for ordinal data
  output$ordinalPlots <- renderPlot({
    data <- parsedData()
    if (is.null(data) || data$type != "ordinal") return(NULL)
    
    values <- data$values
    freq_table <- table(values)
    
    if (length(freq_table) == 0) return(NULL)
    
    # Bar plot for ordinal data
    barplot(freq_table, main = "Ordinal Data Bar Plot", 
            col = "#6BAED6", las = 2, cex.names = 0.8)
  })
  
  # Download handler for descriptive statistics
  output$downloadDescSteps <- downloadHandler(
    filename = function() {
      paste("descriptive_statistics_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      withProgress(message = "Generating report", value = 0, {
        tryCatch({
          doc <- officer::read_docx()
          
          data <- parsedData()
          if (is.null(data)) {
            doc <- doc %>%
              officer::body_add_par("Descriptive Statistics Report", style = "heading 1") %>%
              officer::body_add_par("No data provided for analysis.", style = "Normal")
            print(doc, target = file)
            return()
          }
          
          # Title and basic info
          doc <- doc %>%
            officer::body_add_par("Descriptive Statistics Report", style = "heading 1") %>%
            officer::body_add_par(paste("Generated on:", Sys.Date()), style = "Normal") %>%
            officer::body_add_par(paste("Data type:", data$type), style = "Normal") %>%
            officer::body_add_par(paste("Sample size:", length(data$values)), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # Data preview
          doc <- doc %>%
            officer::body_add_par("Data Preview:", style = "heading 2")
          
          if (length(data$values) <= 20) {
            data_preview <- paste(data$values, collapse = ", ")
            doc <- doc %>% officer::body_add_par(data_preview, style = "Normal")
          } else {
            data_preview <- paste(c(data$values[1:20], "..."), collapse = ", ")
            doc <- doc %>% officer::body_add_par(data_preview, style = "Normal")
          }
          
          doc <- doc %>% officer::body_add_par("", style = "Normal")
          
          # Numerical data analysis
          if (data$type == "numerical") {
            values <- as.numeric(data$values)
            values <- values[!is.na(values)]
            
            if (length(values) > 0) {
              doc <- doc %>%
                officer::body_add_par("Numerical Summary:", style = "heading 2")
              
              summary_data <- data.frame(
                Statistic = c("Mean", "Median", "Standard Deviation", "Variance", 
                              "Minimum", "Maximum", "Range", "Sample Size"),
                Value = c(round(mean(values), 4), round(median(values), 4),
                          round(sd(values), 4), round(var(values), 4),
                          round(min(values), 4), round(max(values), 4),
                          round(max(values) - min(values), 4), length(values))
              )
              
              ft <- flextable::flextable(summary_data) %>%
                flextable::theme_box() %>%
                flextable::autofit()
              
              doc <- flextable::body_add_flextable(doc, ft) %>%
                officer::body_add_par("", style = "Normal")
            }
          }
          
          # Categorical data analysis
          if (data$type %in% c("categorical", "nominal", "ordinal")) {
            freq_table <- table(data$values)
            prop_table <- prop.table(freq_table)
            
            doc <- doc %>%
              officer::body_add_par("Frequency Distribution:", style = "heading 2")
            
            freq_data <- data.frame(
              Category = names(freq_table),
              Frequency = as.numeric(freq_table),
              Percentage = paste0(round(as.numeric(prop_table) * 100, 2), "%")
            )
            
            ft <- flextable::flextable(freq_data) %>%
              flextable::theme_box() %>%
              flextable::autofit()
            
            doc <- flextable::body_add_flextable(doc, ft) %>%
              officer::body_add_par("", style = "Normal")
          }
          
          # Save document
          print(doc, target = file)
          
        }, error = function(e) {
          showNotification(paste("Error generating report:", e$message), type = "error")
        })
      })
    }
  )
  
  # Flowchart generation for Other Formulas
  generateDotCode_other <- function() {
    if (rv_other$stratumCount <= 1) return("")
    
    strata <- sapply(1:rv_other$stratumCount, function(i) input[[paste0("stratum_other", i)]] )
    pops <- sapply(1:rv_other$stratumCount, function(i) input[[paste0("pop_other", i)]] )
    alloc <- allocationData_other()
    total_sample <- sum(alloc$Proportional_Sample[1:rv_other$stratumCount])
    
    nodes <- paste0("node", 1:(rv_other$stratumCount + 3))
    
    # Create node labels based on user preferences
    node_labels <- c(paste0("Total Population\nN = ", totalPopulation_other()))
    
    # Add stratum nodes with optional content
    for (i in 1:rv_other$stratumCount) {
      stratum_label <- paste0("Stratum: ", strata[i])
      if (input$includePopSize_other) {
        stratum_label <- paste0(stratum_label, "\nPop: ", pops[i])
      }
      if (input$includeStratumSize_other) {
        stratum_label <- paste0(stratum_label, "\nSample: ", alloc$Proportional_Sample[i])
      }
      node_labels <- c(node_labels, stratum_label)
    }
    
    node_labels <- c(node_labels, paste0("Total Sample\nn = ", total_sample))
    
    for (i in seq_along(nodes)) {
      if (!is.null(rv_other$nodeTexts[[nodes[i]]])) {
        node_labels[i] <- rv_other$nodeTexts[[nodes[i]]]
      }
    }
    
    edge_labels <- sapply(1:rv_other$stratumCount, function(i) {
      edge_name <- paste0("edge", i)
      if (!is.null(rv_other$edgeTexts[[edge_name]])) {
        return(rv_other$edgeTexts[[edge_name]])
      } else {
        return(paste0("", alloc$Proportional_Sample[i]))
      }
    })
    
    # Get font sizes for nodes and edges
    node_font_sizes <- sapply(nodes, function(n) {
      rv_other$nodeFontSizes[[n]] %||% input$nodeFontSize_other
    })
    
    edge_font_sizes <- sapply(paste0("edge", 1:rv_other$stratumCount), function(e) {
      rv_other$edgeFontSizes[[e]] %||% input$edgeFontSize_other
    })
    
    dot_code <- paste0(
      "digraph flowchart {
        rankdir=TB;
        layout=\"", input$flowchartLayout_other, "\";
        node [fontname=Arial, shape=\"", input$nodeShape_other, "\", style=filled, fillcolor='", input$nodeColor_other, "', 
              width=", input$nodeWidth_other, ", height=", input$nodeHeight_other, "];
        edge [color='", input$edgeColor_other, "', arrowsize=", input$arrowSize_other, "];
        
        // Nodes with custom font sizes
        '", nodes[1], "' [label='", node_labels[1], "', id='", nodes[1], "', fontsize=", node_font_sizes[1], "];
        '", nodes[length(nodes)], "' [label='", node_labels[length(node_labels)], "', id='", nodes[length(nodes)], "', fontsize=", node_font_sizes[length(nodes)], "];
      ",
      paste0(sapply(1:rv_other$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' [label='", node_labels[i+1], "', id='", nodes[i+1], "', fontsize=", node_font_sizes[i+1], "];")
      }), collapse = "\n"),
      "
        
        // Edges with custom font sizes
        '", nodes[1], "' -> {",
      paste0("'", nodes[2:(rv_other$stratumCount+1)], "'", collapse = " "),
      "} [label='', id='stratification_edges', fontsize=", input$edgeFontSize_other, "];
      ",
      paste0(sapply(1:rv_other$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' -> '", nodes[length(nodes)], "' [label='", edge_labels[i], "', id='edge", i, "', fontsize=", edge_font_sizes[i], "];")
      }), collapse = "\n"),
      "
      }"
    )
    
    return(dot_code)
  }
  
  output$flowchart_other <- renderGrViz({
    if (!input$showFlowchart_other || rv_other$stratumCount <= 1) return()
    
    dot_code <- generateDotCode_other()
    grViz(dot_code)
  })
  
  # JavaScript for handling clicks on nodes and edges for Other Formulas
  jsCode_other <- '
  $(document).ready(function() {
    document.getElementById("flowchart_other").addEventListener("click", function(event) {
      var target = event.target;
      if (target.tagName === "text") {
        target = target.parentNode;
      }
      
      if (target.getAttribute("class") && target.getAttribute("class").includes("node")) {
        var nodeId = target.getAttribute("id");
        Shiny.setInputValue("selected_node_other", nodeId);
      } else if (target.getAttribute("class") && target.getAttribute("class").includes("edge")) {
        var pathId = target.getAttribute("id");
        Shiny.setInputValue("selected_edge_other", pathId);
      }
    });
  });
  '
  
  observe({
    session$sendCustomMessage(type='jsCode', list(value = jsCode_other))
  })
  
  # Similar event handlers for Other Formulas flowchart editing
  # (Following the same pattern as previous sections)
  
  observeEvent(input$selected_node_other, {
    rv_other$selectedNode <- input$selected_node_other
    rv_other$selectedEdge <- NULL
    
    if (!is.null(rv_other$nodeTexts[[rv_other$selectedNode]])) {
      updateTextInput(session, "nodeText_other", value = rv_other$nodeTexts[[rv_other$selectedNode]])
    } else {
      node_index <- as.numeric(gsub("node", "", rv_other$selectedNode))
      if (node_index == 1) {
        updateTextInput(session, "nodeText_other", value = paste0("Total Population\nN = ", totalPopulation_other()))
      } else if (node_index == rv_other$stratumCount + 2) {
        alloc <- allocationData_other()
        total_sample <- sum(alloc$Proportional_Sample[1:rv_other$stratumCount])
        updateTextInput(session, "nodeText_other", 
                        value = paste0("Total Sample\nn = ", total_sample))
      } else {
        stratum_num <- node_index - 1
        strata <- sapply(1:rv_other$stratumCount, function(i) input[[paste0("stratum_other", i)]])
        pops <- sapply(1:rv_other$stratumCount, function(i) input[[paste0("pop_other", i)]])
        alloc <- allocationData_other()
        updateTextInput(session, "nodeText_other", 
                        value = paste0("Stratum: ", strata[stratum_num], "\nPop: ", pops[stratum_num], "\nSample: ", alloc$Proportional_Sample[stratum_num]))
      }
    }
    
    if (!is.null(rv_other$nodeFontSizes[[rv_other$selectedNode]])) {
      updateSliderInput(session, "selectedNodeFontSize_other", 
                        value = rv_other$nodeFontSizes[[rv_other$selectedNode]])
    } else {
      updateSliderInput(session, "selectedNodeFontSize_other", value = input$nodeFontSize_other)
    }
  })
  
  observeEvent(input$selected_edge_other, {
    rv_other$selectedEdge <- input$selected_edge_other
    rv_other$selectedNode <- NULL
    
    if (!is.null(rv_other$edgeTexts[[rv_other$selectedEdge]])) {
      updateTextInput(session, "nodeText_other", value = rv_other$edgeTexts[[rv_other$selectedEdge]])
    } else {
      edge_num <- as.numeric(gsub("edge", "", rv_other$selectedEdge))
      alloc <- allocationData_other()
      updateTextInput(session, "nodeText_other", 
                      value = paste0("", alloc$Proportional_Sample[edge_num]))
    }
    
    if (!is.null(rv_other$edgeFontSizes[[rv_other$selectedEdge]])) {
      updateSliderInput(session, "selectedEdgeFontSize_other", 
                        value = rv_other$edgeFontSizes[[rv_other$selectedEdge]])
    } else {
      updateSliderInput(session, "selectedEdgeFontSize_other", value = input$edgeFontSize_other)
    }
  })
  
  # Download handlers for Other Formulas
  output$downloadOtherWord <- downloadHandler(
    filename = function() {
      paste0("other_formulas_results_", Sys.Date(), ".docx")
    },
    content = function(file) {
      withProgress(message = "Generating Word document", value = 0, {
        tryCatch({
          doc <- officer::read_docx()
          
          doc <- doc %>% 
            officer::body_add_par(paste("Sample Size Results -", input$formula_type), style = "heading 1") %>%
            officer::body_add_par(paste("Generated on:", Sys.Date()), style = "Normal") %>%
            officer::body_add_par("", style = "Normal")
          
          # Add formula-specific parameters
          doc <- doc %>%
            officer::body_add_par("Input Parameters:", style = "heading 2")
          
          # Add interpretation
          interp <- interpretationText_other()
          if (is.list(interp)) interp <- paste(unlist(interp), collapse = " ")
          
          doc <- doc %>%
            officer::body_add_par("Interpretation:", style = "heading 2") %>%
            officer::body_add_par(as.character(interp), style = "Normal")
          
          # Add allocation table if strata exist
          if (rv_other$stratumCount > 1) {
            alloc_data <- allocationData_other()
            if (!is.null(alloc_data)) {
              doc <- doc %>%
                officer::body_add_par("Allocation Results:", style = "heading 2")
              
              ft_alloc <- flextable::flextable(alloc_data) %>%
                flextable::theme_box() %>%
                flextable::autofit()
              doc <- flextable::body_add_flextable(doc, ft_alloc)
            }
          }
          
          print(doc, target = file)
          
        }, error = function(e) {
          showNotification(paste("Error generating document:", e$message), type = "error")
        })
      })
    }
  )
  
  output$downloadOtherSteps <- downloadHandler(
    filename = function() {
      paste("other_formulas_calculation_steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      steps <- c(
        paste("SAMPLE SIZE CALCULATION -", toupper(input$formula_type)),
        "==========================================",
        paste("Date:", Sys.Date()),
        "",
        "INPUT PARAMETERS:",
        paste("Formula type:", input$formula_type),
        paste("Non-response rate:", input$non_response_other, "%"),
        "",
        "SAMPLE SIZE CALCULATION:",
        paste("Required sample size:", otherSampleSize()),
        paste("Adjusted for non-response:", adjustedSampleSize_other()),
        "",
        "STRATUM INFORMATION:"
      )
      
      for (i in 1:rv_other$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", input[[paste0("stratum_other", i)]], 
                         "- Population:", input[[paste0("pop_other", i)]]))
      }
      
      if (rv_other$stratumCount > 1) {
        steps <- c(steps,
                   paste("Total population:", totalPopulation_other()),
                   "",
                   "PROPORTIONAL ALLOCATION CALCULATIONS:",
                   "Formula: n_i = (N_i / N) × n"
        )
        
        alloc <- allocationData_other()
        for (i in 1:rv_other$stratumCount) {
          steps <- c(steps, 
                     paste("Stratum", i, ":", alloc$Calculation[i]))
        }
        
        steps <- c(steps,
                   "",
                   "FINAL ALLOCATION:"
        )
        
        for (i in 1:rv_other$stratumCount) {
          steps <- c(steps, 
                     paste("Stratum", i, ":", alloc$Proportional_Sample[i], "samples"))
        }
        
        steps <- c(steps,
                   paste("Total samples:", sum(alloc$Proportional_Sample[1:rv_other$stratumCount]))
        )
      }
      
      steps <- c(steps,
                 "",
                 "INTERPRETATION:",
                 interpretationText_other()
      )
      
      writeLines(steps, file)
    }
  )
  
  # Add similar download handlers for flowchart exports in Other Formulas section
  output$downloadFlowchartPNG_other <- downloadHandler(
    filename = function() {
      paste("other_formulas_flowchart_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_other()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_png(tmp_file, file, width = 3000, height = 2000)
      unlink(tmp_file)
    }
  )
  
  output$downloadFlowchartSVG_other <- downloadHandler(
    filename = function() {
      paste("other_formulas_flowchart_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_other()
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(file)
    }
  )
  
  output$downloadFlowchartPDF_other <- downloadHandler(
    filename = function() {
      paste("other_formulas_flowchart_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_other()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_pdf(tmp_file, file)
      unlink(tmp_file)
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
