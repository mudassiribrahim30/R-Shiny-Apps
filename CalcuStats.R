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

getGreeting <- function() {
  hour <- as.numeric(format(Sys.time(), "%H"))
  if (hour < 12) {
    return("Good morning!")
  } else if (hour < 18) {
    return("Good afternoon!")
  } else {
    return("Good evening!")
  }
}

# WHO-inspired CSS
who_css <- "
/* WHO-inspired color scheme */
:root {
  --who-blue: #0092D0;
  --who-light-blue: #6BC1E0;
  --who-dark-blue: #00689D;
  --who-green: #7CC242;
  --who-gray: #6D6E71;
  --who-light-gray: #F1F1F2;
}

/* WHO-style header and navigation */
.navbar {
  background-color: var(--who-blue) !important;
  border: none;
  border-radius: 0;
  margin-bottom: 0;
}

.navbar .navbar-nav > li > a {
  color: white !important;
  font-weight: 500;
  font-size: 15px;
}

.navbar .navbar-nav > li > a:hover {
  background-color: var(--who-dark-blue) !important;
  color: white !important;
}

.navbar .navbar-brand {
  color: white !important;
  font-weight: bold;
  font-size: 18px;
}

/* WHO-style panels and cards */
.panel {
  border: none;
  border-radius: 4px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.panel-heading {
  background-color: var(--who-blue) !important;
  color: white !important;
  border: none;
  border-radius: 4px 4px 0 0;
  font-weight: bold;
}

.who-info-box {
  background-color: #E8F4FC;
  border-left: 4px solid var(--who-blue);
  padding: 15px;
  margin-bottom: 20px;
  border-radius: 4px;
}

.who-formula-box {
  background-color: #F8F9FA;
  border: 1px solid #DEE2E6;
  padding: 15px;
  margin-bottom: 20px;
  border-radius: 4px;
  font-family: 'Courier New', monospace;
}

/* WHO-style buttons */
.btn-who {
  background-color: var(--who-blue);
  color: white;
  border: none;
}

.btn-who:hover {
  background-color: var(--who-dark-blue);
  color: white;
}

.btn-who-secondary {
  background-color: var(--who-green);
  color: white;
  border: none;
}

.btn-who-secondary:hover {
  background-color: #6BA83A;
  color: white;
}

/* WHO-style footer */
.who-footer {
  background-color: var(--who-gray);
  color: white;
  padding: 15px 0;
  text-align: center;
  margin-top: 30px;
}

.who-footer a {
  color: var(--who-light-blue);
  text-decoration: underline;
}

/* Improved form controls */
.form-control {
  border-radius: 4px;
  border: 1px solid #CED4DA;
}

.form-control:focus {
  border-color: var(--who-light-blue);
  box-shadow: 0 0 0 0.2rem rgba(0, 146, 208, 0.25);
}

/* Calculation sections - WHITE BACKGROUND WITH BLACK TEXT - LARGER FONT */
.shiny-text-output, 
.verbatimTextOutput,
pre {
  background-color: white !important;
  color: black !important;
  padding: 20px !important;
  border-radius: 6px !important;
  border: 2px solid #E8F4FC !important;
  font-family: 'Arial', sans-serif !important;  /* Changed to Arial for better readability */
  font-size: 18px !important;  /* Increased font size */
  line-height: 1.6 !important;
  margin-bottom: 20px !important;
  white-space: pre-wrap !important;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
}

/* Specific styling for calculation steps */
.shiny-text-output pre,
.verbatimTextOutput pre {
  font-size: 18px !important;  /* Increased font size */
  line-height: 1.8 !important;
  margin: 15px 0 !important;
  padding: 20px !important;
  background-color: #ffffff !important;
  border-left: 4px solid var(--who-blue) !important;
}

/* Table styling for calculation results */
.table-responsive table {
  background-color: white !important;
  color: black !important;
  border: 2px solid #E8F4FC !important;
  font-size: 18px !important;  /* Increased font size */
  margin: 20px 0 !important;
}

.table-responsive th {
  background-color: var(--who-light-blue) !important;
  color: black !important;
  font-size: 18px !important;  /* Increased font size */
  font-weight: bold !important;
  padding: 15px !important;
  text-align: center !important;
}

.table-responsive td {
  background-color: white !important;
  color: black !important;
  border: 1px solid #ddd !important;
  font-size: 18px !important;  /* Increased font size */
  padding: 12px 15px !important;
}

/* Headings above calculation sections */
h3, h4 {
  color: var(--who-dark-blue) !important;
  margin-top: 25px !important;
  margin-bottom: 15px !important;
}

h3 {
  font-size: 26px !important;  /* Increased font size */
  font-weight: bold !important;
}

h4 {
  font-size: 22px !important;  /* Increased font size */
  font-weight: 600 !important;
}

/* Fluid row spacing for calculation outputs */
.fluid-row {
  margin-bottom: 25px !important;
}

.fluid-row .col-md-6 {
  margin-bottom: 20px !important;
}

/* Welcome header and date styling */
.welcome-header {
  text-align: right;
  padding: 10px 20px !important;  /* Increased padding */
  background-color: #f8f9fa;
  font-size: 16px !important;  /* Increased font size */
  color: #00689D;
  border-bottom: 1px solid #E8F4FC;
  font-family: 'Arial', sans-serif !important;  /* Added font family */
}

.welcome-header div {
  display: inline-block;
  margin-right: 15px !important;  /* Increased margin */
}

.welcome-header div:last-child {
  border-left: 1px solid #ccc;
  padding-left: 15px !important;  /* Increased padding */
}

/* Responsive adjustments for better mobile viewing */
@media (max-width: 768px) {
  .sidebar-panel {
    margin-bottom: 20px;
  }
  
  .shiny-text-output, 
  .verbatimTextOutput,
  pre {
    padding: 15px !important;
    font-size: 16px !important;  /* Adjusted for mobile */
    line-height: 1.5 !important;
    margin-bottom: 15px !important;
  }
  
  .table-responsive th,
  .table-responsive td {
    font-size: 16px !important;  /* Adjusted for mobile */
    padding: 10px !important;
  }
  
  h3 {
    font-size: 22px !important;  /* Adjusted for mobile */
  }
  
  h4 {
    font-size: 20px !important;  /* Adjusted for mobile */
  }
  
  .welcome-header {
    font-size: 14px !important;  /* Adjusted for mobile */
    padding: 8px 15px !important;
  }
}

/* Additional spacing for better readability */
.container-fluid {
  padding: 0 20px !important;
}

.main-panel {
  padding: 20px !important;
}

/* Highlight important numbers in calculations */
strong {
  color: var(--who-dark-blue) !important;
  font-weight: bold !important;
}

em {
  color: #2C3E50 !important;
  font-style: italic !important;
}
"

ui <- navbarPage(
  title = div(icon("chart-bar"), "CalcuStats"),
  theme = shinytheme("flatly"),
  header = tags$head(
    tags$style(HTML(who_css)),
    tags$div(
      class = "welcome-header",
      div(style = "display: inline-block; margin-right: 15px;", 
          textOutput("welcomeMessage")),
      div(style = "display: inline-block; border-left: 1px solid #ccc; padding-left: 15px;", 
          textOutput("currentDateTime"))
    )
  ),
  # ... rest of the UI code remains the same
  footer = div(
    class = "who-footer",
    HTML("<p>© 2025 Mudasir Mohammed Ibrahim. All rights reserved. | 
         <a href='https://github.com/mudassiribrahim30' target='_blank'>GitHub Profile</a></p>"),
    div(
      style = "margin-top: 10px; font-size: 0.9em; color: #fff;",
      "Your Companion for Sample Size Calculation and Descriptive Analytics"
    )
  ),
  useShinyjs(),
  
  # Home/Introduction Tab
  tabPanel("Home",
           div(class = "container-fluid",
               div(class = "row",
                   div(class = "col-md-8 col-md-offset-2",
                       div(class = "jumbotron", style = "background-color: #E8F4FC; padding: 30px;",
                           h1("Welcome to CalcuStats", style = "color: #00689D;"),
                           p("A comprehensive statistical tool for sample size calculation, power analysis, and descriptive statistics."),
                           hr(),
                           p("This application provides researchers and healthcare professionals with reliable methods for determining appropriate sample sizes for various study designs."),
                           p("Developed by Mudasir Mohammed Ibrahim, BSc, RN")
                       ),
                       
                       div(class = "row",
                           div(class = "col-md-4",
                               div(class = "panel panel-default",
                                   div(class = "panel-heading", 
                                       h3("Sample Size Calculators", class = "panel-title")
                                   ),
                                   div(class = "panel-body",
                                       p("Determine appropriate sample sizes using:"),
                                       tags$ul(
                                         tags$li("Taro Yamane formula"),
                                         tags$li("Cochran's formula"),
                                         tags$li("Proportional allocation"),
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
                                       p("Calculate statistical power for:"),
                                       tags$ul(
                                         tags$li("t-tests"),
                                         tags$li("ANOVA"),
                                         tags$li("Correlation studies"),
                                         tags$li("Regression analysis"),
                                         tags$li("Chi-square tests")
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
                                         tags$li("Data visualization options")
                                       )
                                   )
                               )
                           )
                       ),
                       
                       div(class = "who-info-box",
                           h4("Getting Started"),
                           p("1. Select the appropriate calculator from the navigation menu"),
                           p("2. Enter your study parameters"),
                           p("3. Review the results and download reports"),
                           p("4. Use the flowchart features to visualize your sampling strategy")
                       )
                   )
               )
           )
  ),
  
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
                     numericInput("p", "Estimated Proportion (p)", value = 0.5, min = 0.01, max = 0.99, step = 0.01),
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
  
  # REVISED: Other Formulas Tab
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
                                 choices = c("Mean Estimation (Known Variance)" = "mean_known_var",
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
                     style = "margin-top: 20px; border: 1px solid #ddd; padding: 10px; font-size: 14px;",
                     h4("Interactive Diagram Editor"),
                     p("Click on nodes or edges in the diagram to edit them."),
                     uiOutput("nodeEdgeEditorUI_other")
                 )
               )
             )
           )
  ),
  
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
                     numericInput("effectSize", "Effect Size", value = 0.5),
                     numericInput("alpha", "Significance Level (alpha)", value = 0.05),
                     numericInput("power", "Desired Power", value = 0.8),
                     conditionalPanel(
                       condition = "input.testType == 'Multiple Linear Regression'",
                       numericInput("predictors", "Number of Predictors", value = 2, min = 1)
                     ),
                     actionButton("runPower", "Run Power Analysis", class = "btn-who")
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
            <p>A power of 0.8 (80%) is generally considered acceptable in most research contexts.</p>
          ")
               ),
               div(style = "font-size: 14px;", verbatimTextOutput("powerResult"))
             )
           )
  ),
  
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
  
  # User Guide Tab (now properly maintained as its own tab)
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
               p("With expertise in both healthcare and data analysis, Mudasir created this tool to help researchers and students perform essential statistical calculations with ease."),
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
  )
)


server <- function(input, output, session) {
  # Welcome message and date/time
  output$welcomeMessage <- renderText({
    paste(getGreeting(), "Welcome to CalcuStats!")
  })
  
  output$currentDateTime <- renderText({
    format(Sys.time(), "%A, %B %d, %Y %I:%M %p")
  })
  
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
      p = 0.5,
      z = 1.96,
      e_c = 0.05,
      N_c = NULL,
      non_response_c = 0,
      stratumCount = 1,
      stratum_names = c("Stratum 1"),
      stratum_pops = c(100)
    ),
    other = list(
      formula_type = "mean_known_var",
      alpha_other = 0.05,
      power_other = 0.8,
      effect_size_other = 0.5,
      sigma_other = 1,
      d_other = 0.1,
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
        updateNumericInput(session, "p", value = cochran_vals$p)
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
      p = input$p,
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
    updateNumericInput(session, "p", value = initValues$cochran$p)
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
  
  cochranSampleSize <- reactive({
    p <- input$p
    z <- input$z
    e <- input$e_c
    N_c <- input$N_c
    
    if (is.null(p) || is.null(z) || is.null(e) || e == 0) return(0)
    
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
    p <- input$p
    z <- input$z
    e <- input$e_c
    N_c <- input$N_c
    n0 <- (z^2 * p * (1 - p)) / (e^2)
    
    explanation <- paste(
      "Cochran's Formula Calculation:\n",
      "1. Infinite population formula: n₀ = (Z² × p × (1-p)) / e²\n",
      "2. Calculation: (", z, "² × ", p, " × (1-", p, ")) / ", e, "²\n",
      "3. Step-by-step: (", z^2, " × ", p, " × ", (1-p), ") / ", e^2, "\n",
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
    paste0("Based on Cochran's formula with p = ", input$p, ", Z = ", input$z,
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
            officer::body_add_par(paste("Estimated proportion (p):", input$p), style = "Normal") %>%
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
        paste("Estimated proportion (p):", input$p),
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
                 paste("p × (1-p) = ", input$p, " × (1-", input$p, ") = ", input$p * (1-input$p)),
                 paste("Z² × p × (1-p) = ", input$z^2, " × ", input$p * (1-input$p), " = ", input$z^2 * input$p * (1-input$p)),
                 paste("e² = ", input$e_c, "² = ", input$e_c^2),
                 paste("n₀ = ", input$z^2 * input$p * (1-input$p), " / ", input$e_c^2, " = ", 
                       (input$z^2 * input$p * (1-input$p)) / (input$e_c^2))
      )
      
      if (!is.null(input$N_c) && !is.na(input$N_c) && input$N_c > 0) {
        n0 <- (input$z^2 * input$p * (1-input$p)) / (input$e_c^2)
        steps <- c(steps,
                   "",
                   "FINITE POPULATION CORRECTION:",
                   "n = n₀ / (1 + (n₀ - 1)/N)",
                   paste("Calculation: ", n0, " / (1 + (", n0, " - 1)/", input$N_c, ")"),
                   paste("Step-by-step: ", n0, " / (1 + ", (n0 - 1), "/", input$N_c, ") = ", 
                         n0, " / (1 + ", (n0 - 1)/input$N_c, ")"),
                   paste("Result: ", n0, " / ", (1 + (n0 - 1)/input$N_c), " = ", 
                         n0 / (1 + (n0 - 1)/input$N_c))
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
  
  output$formula_params <- renderUI({
    formula_type <- input$formula_type
    
    params <- switch(formula_type,
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
                       numericInput("p1_other", "Proportion 1 (p₁)", value = 0.5, min = 0.01, max = 0.99, step = 0.01),
                       numericInput("p2_other", "Proportion 2 (p₂)", value = 0.3, min = 0.01, max = 0.99, step = 0.01)
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
                       numericInput("p1_other", "Proportion in Control Group", value = 0.2, min = 0.01, max = 0.99, step = 0.01)
                     ),
                     "relative_risk" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("rr_other", "Relative Risk (RR)", value = 1.5, min = 1.1, step = 0.1),
                       numericInput("p1_other", "Proportion in Control Group", value = 0.2, min = 0.01, max = 0.99, step = 0.01)
                     ),
                     "prevalence" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("precision_other", "Precision (d)", value = 0.05, min = 0.01, step = 0.01),
                       numericInput("prevalence_other", "Expected Prevalence", value = 0.1, min = 0.01, max = 0.99, step = 0.01)
                     ),
                     "case_control" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("or_other", "Odds Ratio (OR)", value = 2.0, min = 1.1, step = 0.1),
                       numericInput("p1_other", "Proportion in Controls", value = 0.2, min = 0.01, max = 0.99, step = 0.01),
                       numericInput("r_other", "Case:Control Ratio", value = 1, min = 0.1, step = 0.1)
                     ),
                     "cohort" = tagList(
                       numericInput("alpha_other", "Alpha (α)", value = 0.05, min = 0.001, max = 0.2, step = 0.01),
                       numericInput("power_other", "Power (1-β)", value = 0.8, min = 0.5, max = 0.99, step = 0.01),
                       numericInput("rr_other", "Relative Risk (RR)", value = 1.5, min = 1.1, step = 0.1),
                       numericInput("p1_other", "Proportion in Unexposed", value = 0.2, min = 0.01, max = 0.99, step = 0.01),
                       numericInput("r_other", "Exposed:Unexposed Ratio", value = 1, min = 0.1, step = 0.1)
                     )
    )
    
    return(params)
  })
  
  output$formula_description <- renderUI({
    formula_type <- input$formula_type
    
    description <- switch(formula_type,
                          "mean_known_var" = div(
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
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
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
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
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
                            HTML("
                <strong>Difference Between Two Proportions:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (p₁ - p₂)²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>p₁, p₂</strong> = proportions in two groups</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For comparing two independent proportions.
              ")
                          ),
                          "correlation" = div(
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
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
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
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
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
                            HTML("
                <strong>Odds Ratio:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (ln(OR))²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>OR</strong> = odds ratio</li>
                  <li><strong>p₁</strong> = proportion in control group</li>
                  <li><strong>p₂</strong> = p₁ × OR / (1 + p₁(OR-1))</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For case-control studies testing odds ratios.
              ")
                          ),
                          "relative_risk" = div(
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
                            HTML("
                <strong>Relative Risk:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (ln(RR))²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>RR</strong> = relative risk</li>
                  <li><strong>p₁</strong> = proportion in control group</li>
                  <li><strong>p₂</strong> = p₁ × RR</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For cohort studies testing relative risk.
              ")
                          ),
                          "prevalence" = div(
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
                            HTML("
                <strong>Prevalence Study:</strong><br>
                <em>n = (Z² × p × (1-p)) / d²</em><br>
                <ul>
                  <li><strong>Z</strong> = Z-score for desired confidence level</li>
                  <li><strong>p</strong> = expected prevalence</li>
                  <li><strong>d</strong> = precision (margin of error)</li>
                </ul>
                For estimating disease prevalence with specified precision.
              ")
                          ),
                          "case_control" = div(
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
                            HTML("
                <strong>Case-Control Study:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (p₁ - p₂)²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>OR</strong> = odds ratio</li>
                  <li><strong>p₁</strong> = proportion in controls</li>
                  <li><strong>p₂</strong> = p₁ × OR / (1 + p₁(OR-1))</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For case-control studies with specified case:control ratio.
              ")
                          ),
                          "cohort" = div(
                            style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
                            HTML("
                <strong>Cohort Study:</strong><br>
                <em>n = [Zα√(2p(1-p)) + Zβ√(p₁(1-p₁) + p₂(1-p₂))]² / (p₁ - p₂)²</em><br>
                <ul>
                  <li><strong>Zα</strong> = Z-score for alpha</li>
                  <li><strong>Zβ</strong> = Z-score for beta (1-power)</li>
                  <li><strong>RR</strong> = relative risk</li>
                  <li><strong>p₁</strong> = proportion in unexposed</li>
                  <li><strong>p₂</strong> = p₁ × RR</li>
                  <li><strong>p</strong> = (p₁ + p₂)/2</li>
                </ul>
                For cohort studies with specified exposed:unexposed ratio.
              ")
                          )
    )
    
    return(description)
  })
  
  otherSampleSize <- reactive({
    formula_type <- input$formula_type
    
    tryCatch({
      result <- switch(formula_type,
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
                         p1 <- input$p1_other
                         p2 <- input$p2_other
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
                         p1 <- input$p1_other
                         p2 <- p1 * or / (1 + p1 * (or - 1))
                         result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                               alternative = "two.sided")
                         ceiling(result$n * 2)  # Total sample size for two groups
                       },
                       "relative_risk" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         rr <- input$rr_other
                         p1 <- input$p1_other
                         p2 <- p1 * rr
                         result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                               alternative = "two.sided")
                         ceiling(result$n * 2)  # Total sample size for two groups
                       },
                       "prevalence" = {
                         alpha <- input$alpha_other
                         p <- input$prevalence_other
                         d <- input$precision_other
                         z <- qnorm(1 - alpha/2)
                         ceiling((z^2 * p * (1 - p)) / (d^2))
                       },
                       "case_control" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         or <- input$or_other
                         p1 <- input$p1_other
                         r <- input$r_other  # case:control ratio
                         p2 <- p1 * or / (1 + p1 * (or - 1))
                         result <- pwr.2p2n.test(h = ES.h(p1, p2), n1 = NULL, n2 = r, 
                                                 sig.level = alpha, power = power, 
                                                 alternative = "two.sided")
                         ceiling(result$n1 + result$n2)  # Total sample size
                       },
                       "cohort" = {
                         alpha <- input$alpha_other
                         power <- input$power_other
                         rr <- input$rr_other
                         p1 <- input$p1_other
                         r <- input$r_other  # exposed:unexposed ratio
                         p2 <- p1 * rr
                         result <- pwr.2p2n.test(h = ES.h(p1, p2), n1 = NULL, n2 = r, 
                                                 sig.level = alpha, power = power, 
                                                 alternative = "two.sided")
                         ceiling(result$n1 + result$n2)  # Total sample size
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
  
  output$formulaExplanation_other <- renderText({
    formula_type <- input$formula_type
    
    explanation <- switch(formula_type,
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
                            p1 <- input$p1_other
                            p2 <- input$p2_other
                            result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                                  alternative = "two.sided")
                            paste(
                              "Difference Between Two Proportions:\n",
                              "Proportions: p1 = ", p1, ", p2 = ", p2, "\n",
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
                            p1 <- input$p1_other
                            p2 <- p1 * or / (1 + p1 * (or - 1))
                            result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                                  alternative = "two.sided")
                            paste(
                              "Odds Ratio:\n",
                              "Odds ratio: ", or, "\n",
                              "Control proportion: ", p1, "\n",
                              "Case proportion: ", round(p2, 3), "\n",
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
                            p1 <- input$p1_other
                            p2 <- p1 * rr
                            result <- pwr.2p.test(h = ES.h(p1, p2), sig.level = alpha, power = power, 
                                                  alternative = "two.sided")
                            paste(
                              "Relative Risk:\n",
                              "Relative risk: ", rr, "\n",
                              "Unexposed proportion: ", p1, "\n",
                              "Exposed proportion: ", round(p2, 3), "\n",
                              "Effect size (h): ", round(ES.h(p1, p2), 3), "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Sample size per group: ", ceiling(result$n), "\n",
                              "Total sample size: ", ceiling(result$n * 2)
                            )
                          },
                          "prevalence" = {
                            alpha <- input$alpha_other
                            p <- input$prevalence_other
                            d <- input$precision_other
                            z <- qnorm(1 - alpha/2)
                            paste(
                              "Prevalence Study:\n",
                              "Formula: n = (Z² × p × (1-p)) / d²\n",
                              "Calculation: (", round(z, 3), "² × ", p, " × (1-", p, ")) / ", d, "²\n",
                              "Step-by-step: (", round(z^2, 3), " × ", p, " × ", (1-p), ") / ", d^2, "\n",
                              "Result: ", (z^2 * p * (1-p)), " / ", d^2, " = ", (z^2 * p * (1-p)) / d^2, "\n",
                              "Apply ceiling function: ", otherSampleSize()
                            )
                          },
                          "case_control" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            or <- input$or_other
                            p1 <- input$p1_other
                            r <- input$r_other
                            p2 <- p1 * or / (1 + p1 * (or - 1))
                            result <- pwr.2p2n.test(h = ES.h(p1, p2), n1 = NULL, n2 = r, 
                                                    sig.level = alpha, power = power, 
                                                    alternative = "two.sided")
                            paste(
                              "Case-Control Study:\n",
                              "Odds ratio: ", or, "\n",
                              "Control proportion: ", p1, "\n",
                              "Case proportion: ", round(p2, 3), "\n",
                              "Case:Control ratio: ", r, "\n",
                              "Effect size (h): ", round(ES.h(p1, p2), 3), "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Cases needed: ", ceiling(result$n1), "\n",
                              "Controls needed: ", ceiling(result$n2), "\n",
                              "Total sample size: ", ceiling(result$n1 + result$n2)
                            )
                          },
                          "cohort" = {
                            alpha <- input$alpha_other
                            power <- input$power_other
                            rr <- input$rr_other
                            p1 <- input$p1_other
                            r <- input$r_other
                            p2 <- p1 * rr
                            result <- pwr.2p2n.test(h = ES.h(p1, p2), n1 = NULL, n2 = r, 
                                                    sig.level = alpha, power = power, 
                                                    alternative = "two.sided")
                            paste(
                              "Cohort Study:\n",
                              "Relative risk: ", rr, "\n",
                              "Unexposed proportion: ", p1, "\n",
                              "Exposed proportion: ", round(p2, 3), "\n",
                              "Exposed:Unexposed ratio: ", r, "\n",
                              "Effect size (h): ", round(ES.h(p1, p2), 3), "\n",
                              "Alpha: ", alpha, ", Power: ", power, "\n",
                              "Exposed needed: ", ceiling(result$n1), "\n",
                              "Unexposed needed: ", ceiling(result$n2), "\n",
                              "Total sample size: ", ceiling(result$n1 + result$n2)
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
  
  # JavaScript for handling clicks on nodes and edges
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
    
    # Update font size slider for selected node
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
    
    # Update font size slider for selected edge
    if (!is.null(rv_other$edgeFontSizes[[rv_other$selectedEdge]])) {
      updateSliderInput(session, "selectedEdgeFontSize_other", 
                        value = rv_other$edgeFontSizes[[rv_other$selectedEdge]])
    } else {
      updateSliderInput(session, "selectedEdgeFontSize_other", value = input$edgeFontSize_other)
    }
  })
  
  observeEvent(input$editNodeText_other, {
    if (input$newText_other == "") return()
    
    if (!is.null(rv_other$selectedNode)) {
      if (grepl("->", input$newText_other)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_other$nodeTexts[[rv_other$selectedNode]] %||% input$nodeText_other
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_other$nodeTexts[[rv_other$selectedNode]] <- updated_text
      } else {
        # Direct text replacement
        rv_other$nodeTexts[[rv_other$selectedNode]] <- input$newText_other
      }
    } else if (!is.null(rv_other$selectedEdge)) {
      if (grepl("->", input$newText_other)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_other$edgeTexts[[rv_other$selectedEdge]] %||% input$nodeText_other
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- updated_text
      } else {
        # Direct text replacement
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- input$newText_other
      }
    }
    
    updateTextInput(session, "newText_other", value = "")
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$deleteText_other, {
    if (!is.null(rv_other$selectedNode)) {
      if (grepl("->", input$newText_other)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_other$nodeTexts[[rv_other$selectedNode]] %||% input$nodeText_other
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_other$nodeTexts[[rv_other$selectedNode]] <- updated_text
      } else {
        # Delete entire text
        rv_other$nodeTexts[[rv_other$selectedNode]] <- NULL
      }
    } else if (!is.null(rv_other$selectedEdge)) {
      if (grepl("->", input$newText_other)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_other$edgeTexts[[rv_other$selectedEdge]] %||% input$nodeText_other
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- updated_text
      } else {
        # Delete entire text
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- NULL
      }
    }
    
    updateTextInput(session, "newText_other", value = "")
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$applyFontSizes_other, {
    if (!is.null(rv_other$selectedNode)) {
      rv_other$nodeFontSizes[[rv_other$selectedNode]] <- input$selectedNodeFontSize_other
    } else if (!is.null(rv_other$selectedEdge)) {
      rv_other$edgeFontSizes[[rv_other$selectedEdge]] <- input$selectedEdgeFontSize_other
    }
    
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetDiagramText_other, {
    rv_other$nodeTexts <- list()
    rv_other$edgeTexts <- list()
    rv_other$nodeFontSizes <- list()
    rv_other$edgeFontSizes <- list()
    updateTextInput(session, "nodeText_other", value = "")
    updateTextInput(session, "newText_other", value = "")
  })
  
  output$flowchart_other <- renderGrViz({
    if (!input$showFlowchart_other || rv_other$stratumCount <= 1) return()
    
    dot_code <- generateDotCode_other()
    grViz(dot_code)
  })
  
  # JavaScript for handling clicks on nodes and edges
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
    
    # Update font size slider for selected node
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
    
    # Update font size slider for selected edge
    if (!is.null(rv_other$edgeFontSizes[[rv_other$selectedEdge]])) {
      updateSliderInput(session, "selectedEdgeFontSize_other", 
                        value = rv_other$edgeFontSizes[[rv_other$selectedEdge]])
    } else {
      updateSliderInput(session, "selectedEdgeFontSize_other", value = input$edgeFontSize_other)
    }
  })
  
  observeEvent(input$editNodeText_other, {
    if (input$newText_other == "") return()
    
    if (!is.null(rv_other$selectedNode)) {
      if (grepl("->", input$newText_other)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_other$nodeTexts[[rv_other$selectedNode]] %||% input$nodeText_other
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_other$nodeTexts[[rv_other$selectedNode]] <- updated_text
      } else {
        # Direct text replacement
        rv_other$nodeTexts[[rv_other$selectedNode]] <- input$newText_other
      }
    } else if (!is.null(rv_other$selectedEdge)) {
      if (grepl("->", input$newText_other)) {
        # Handle replacement pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        old_text <- trimws(parts[1])
        new_text <- trimws(parts[2])
        current_text <- rv_other$edgeTexts[[rv_other$selectedEdge]] %||% input$nodeText_other
        updated_text <- gsub(old_text, new_text, current_text, fixed = TRUE)
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- updated_text
      } else {
        # Direct text replacement
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- input$newText_other
      }
    }
    
    updateTextInput(session, "newText_other", value = "")
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$deleteText_other, {
    if (!is.null(rv_other$selectedNode)) {
      if (grepl("->", input$newText_other)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_other$nodeTexts[[rv_other$selectedNode]] %||% input$nodeText_other
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_other$nodeTexts[[rv_other$selectedNode]] <- updated_text
      } else {
        # Delete entire text
        rv_other$nodeTexts[[rv_other$selectedNode]] <- NULL
      }
    } else if (!is.null(rv_other$selectedEdge)) {
      if (grepl("->", input$newText_other)) {
        # Handle deletion pattern
        parts <- strsplit(input$newText_other, "->")[[1]]
        text_to_delete <- trimws(parts[1])
        current_text <- rv_other$edgeTexts[[rv_other$selectedEdge]] %||% input$nodeText_other
        updated_text <- gsub(text_to_delete, "", current_text, fixed = TRUE)
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- updated_text
      } else {
        # Delete entire text
        rv_other$edgeTexts[[rv_other$selectedEdge]] <- NULL
      }
    }
    
    updateTextInput(session, "newText_other", value = "")
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$applyFontSizes_other, {
    if (!is.null(rv_other$selectedNode)) {
      rv_other$nodeFontSizes[[rv_other$selectedNode]] <- input$selectedNodeFontSize_other
    } else if (!is.null(rv_other$selectedEdge)) {
      rv_other$edgeFontSizes[[rv_other$selectedEdge]] <- input$selectedEdgeFontSize_other
    }
    
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetDiagramText_other, {
    rv_other$nodeTexts <- list()
    rv_other$edgeTexts <- list()
    rv_other$nodeFontSizes <- list()
    rv_other$edgeFontSizes <- list()
    updateTextInput(session, "nodeText_other", value = "")
    updateTextInput(session, "newText_other", value = "")
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  output$nodeEdgeEditorUI_other <- renderUI({
    if (!is.null(rv_other$selectedNode)) {
      tagList(
        p(strong("Editing Node:"), rv_other$selectedNode),
        textInput("nodeText_other", "Node Text:", value = rv_other$nodeTexts[[rv_other$selectedNode]] %||% ""),
        actionButton("editNodeText_other", "Update Node Text", class = "btn-info"),
        actionButton("resetNodeText_other", "Reset This Text", class = "btn-warning")
      )
    } else if (!is.null(rv_other$selectedEdge)) {
      tagList(
        p(strong("Editing Edge:"), rv_other$selectedEdge),
        textInput("nodeText_other", "Edge Label:", value = rv_other$edgeTexts[[rv_other$selectedEdge]] %||% ""),
        actionButton("editNodeText_other", "Update Edge Label", class = "btn-info"),
        actionButton("resetEdgeText_other", "Reset This Text", class = "btn-warning")
      )
    } else {
      p("Click on a node or edge in the diagram to edit it.")
    }
  })
  
  observeEvent(input$resetNodeText_other, {
    rv_other$nodeTexts[[rv_other$selectedNode]] <- NULL
    rv_other$nodeFontSizes[[rv_other$selectedNode]] <- NULL
    updateTextInput(session, "nodeText_other", value = "")
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetEdgeText_other, {
    rv_other$edgeTexts[[rv_other$selectedEdge]] <- NULL
    rv_other$edgeFontSizes[[rv_other$selectedEdge]] <- NULL
    updateTextInput(session, "nodeText_other", value = "")
    output$flowchart_other <- renderGrViz({
      dot_code <- generateDotCode_other()
      grViz(dot_code)
    })
  })
  
  output$downloadFlowchartPNG_other <- downloadHandler(
    filename = function() {
      paste("other_formula_allocation_flowchart_", Sys.Date(), ".png", sep = "")
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
      paste("other_formula_allocation_flowchart_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_other()
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(file)
    }
  )
  
  output$downloadFlowchartPDF_other <- downloadHandler(
    filename = function() {
      paste("other_formula_allocation_flowchart_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_other()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_pdf(tmp_file, file)
      unlink(tmp_file)
    }
  )
  
  output$downloadOtherWord <- downloadHandler(
    filename = function() {
      paste("other_formula_results_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".docx", sep = "")
    },
    content = function(file) {
      progress <- shiny::Progress$new()
      progress$set(message = "Generating Word document", value = 0)
      on.exit(progress$close())
      
      tryCatch({
        doc <- officer::read_docx()
        
        # ---- Title ----
        doc <- doc %>% 
          officer::body_add_par("Other Formula Sample Size Results", style = "heading 1") %>%
          officer::body_add_par(paste("Generated on:", Sys.Date()), style = "Normal") %>%
          officer::body_add_par("", style = "Normal")
        
        # ---- Formula Type ----
        formula_type_name <- switch(input$formula_type,
                                    "mean_known_var"   = "Mean Estimation (Known Variance)",
                                    "mean_unknown_var" = "Mean Estimation (Unknown Variance)",
                                    "proportion_diff"  = "Difference Between Proportions",
                                    "correlation"      = "Correlation Coefficient",
                                    "regression"       = "Regression Coefficient",
                                    "odds_ratio"       = "Odds Ratio",
                                    "relative_risk"    = "Relative Risk",
                                    "prevalence"       = "Prevalence Study",
                                    "case_control"     = "Case-Control Study",
                                    "cohort"           = "Cohort Study"
        )
        
        doc <- doc %>%
          officer::body_add_par(paste("Formula Type:", formula_type_name), style = "heading 2") %>%
          officer::body_add_par("", style = "Normal")
        
        # ---- Input Parameters ----
        doc <- doc %>% officer::body_add_par("Input Parameters:", style = "heading 2")
        
        params_text <- switch(input$formula_type,
                              "mean_known_var" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Standard Deviation (σ):", input$sigma_other),
                                paste("Margin of Error (d):", input$d_other)
                              ),
                              "mean_unknown_var" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Effect Size (Cohen's d):", input$effect_size_other)
                              ),
                              "proportion_diff" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Proportion 1 (p₁):", input$p1_other),
                                paste("Proportion 2 (p₂):", input$p2_other)
                              ),
                              "correlation" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Correlation Coefficient (r):", input$r_other)
                              ),
                              "regression" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Effect Size (f²):", input$effect_size_other),
                                paste("Number of Predictors:", input$predictors_other)
                              ),
                              "odds_ratio" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Odds Ratio (OR):", input$or_other),
                                paste("Proportion in Control Group:", input$p1_other)
                              ),
                              "relative_risk" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Relative Risk (RR):", input$rr_other),
                                paste("Proportion in Control Group:", input$p1_other)
                              ),
                              "prevalence" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Precision (d):", input$precision_other),
                                paste("Expected Prevalence:", input$prevalence_other)
                              ),
                              "case_control" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Odds Ratio (OR):", input$or_other),
                                paste("Proportion in Controls:", input$p1_other),
                                paste("Case:Control Ratio:", input$r_other)
                              ),
                              "cohort" = c(
                                paste("Alpha (α):", input$alpha_other),
                                paste("Power (1-β):", input$power_other),
                                paste("Relative Risk (RR):", input$rr_other),
                                paste("Proportion in Unexposed:", input$p1_other),
                                paste("Exposed:Unexposed Ratio:", input$r_other)
                              )
        )
        
        for (param in params_text) {
          doc <- doc %>% officer::body_add_par(param, style = "Normal")
        }
        
        doc <- doc %>%
          officer::body_add_par(paste("Non-response rate:", input$non_response_other, "%"), style = "Normal") %>%
          officer::body_add_par(paste("Sample size:", otherSampleSize()), style = "Normal") %>%
          officer::body_add_par(paste("Adjusted sample size:", adjustedSampleSize_other()), style = "Normal") %>%
          officer::body_add_par("", style = "Normal")
        
        progress$set(value = 0.3, detail = "Adding stratum information")
        
        # ---- Stratum Info ----
        doc <- doc %>%
          officer::body_add_par("Stratum Information:", style = "heading 2")
        
        strata_data <- data.frame(
          Stratum = sapply(1:rv_other$stratumCount, function(i) input[[paste0("stratum_other", i)]]),
          Population = sapply(1:rv_other$stratumCount, function(i) input[[paste0("pop_other", i)]])
        )
        
        ft <- flextable::flextable(strata_data) %>%
          flextable::theme_box() %>%
          flextable::autofit()
        doc <- flextable::body_add_flextable(doc, ft) %>%
          officer::body_add_par("", style = "Normal")
        
        progress$set(value = 0.6, detail = "Adding allocation results")
        
        # ---- Allocation Results ----
        doc <- doc %>%
          officer::body_add_par("Allocation Results:", style = "heading 2")
        
        alloc_data <- allocationData_other()
        ft_alloc <- flextable::flextable(alloc_data) %>%
          flextable::theme_box() %>%
          flextable::autofit()
        doc <- flextable::body_add_flextable(doc, ft_alloc) %>%
          officer::body_add_par("", style = "Normal")
        
        progress$set(value = 0.9, detail = "Finalizing document")
        
        # ---- Interpretation ----
        doc <- doc %>%
          officer::body_add_par("Interpretation:", style = "heading 2") %>%
          officer::body_add_par(interpretationText_other(), style = "Normal")
        
        # Save file
        print(doc, target = file)
        
      }, error = function(e) {
        showNotification(paste("Error generating document:", e$message), type = "error")
      })
    }
  )
  
  output$downloadOtherSteps <- downloadHandler(
    filename = function() {
      paste("other_formula_calculation_steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      formula_type_name <- switch(input$formula_type,
                                  "mean_known_var" = "Mean Estimation (Known Variance)",
                                  "mean_unknown_var" = "Mean Estimation (Unknown Variance)",
                                  "proportion_diff" = "Difference Between Proportions",
                                  "correlation" = "Correlation Coefficient",
                                  "regression" = "Regression Coefficient",
                                  "odds_ratio" = "Odds Ratio",
                                  "relative_risk" = "Relative Risk",
                                  "prevalence" = "Prevalence Study",
                                  "case_control" = "Case-Control Study",
                                  "cohort" = "Cohort Study")
      
      steps <- c(
        paste("OTHER FORMULA SAMPLE SIZE CALCULATION STEPS -", formula_type_name),
        "================================================================",
        paste("Date:", Sys.Date()),
        "",
        "INPUT PARAMETERS:"
      )
      
      # Add specific parameters based on formula type
      params_text <- switch(input$formula_type,
                            "mean_known_var" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Standard Deviation (σ):", input$sigma_other),
                              paste("Margin of Error (d):", input$d_other)
                            ),
                            "mean_unknown_var" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Effect Size (Cohen's d):", input$effect_size_other)
                            ),
                            "proportion_diff" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Proportion 1 (p₁):", input$p1_other),
                              paste("Proportion 2 (p₂):", input$p2_other)
                            ),
                            "correlation" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Correlation Coefficient (r):", input$r_other)
                            ),
                            "regression" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Effect Size (f²):", input$effect_size_other),
                              paste("Number of Predictors:", input$predictors_other)
                            ),
                            "odds_ratio" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Odds Ratio (OR):", input$or_other),
                              paste("Proportion in Control Group:", input$p1_other)
                            ),
                            "relative_risk" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Relative Risk (RR):", input$rr_other),
                              paste("Proportion in Control Group:", input$p1_other)
                            ),
                            "prevalence" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Precision (d):", input$precision_other),
                              paste("Expected Prevalence:", input$prevalence_other)
                            ),
                            "case_control" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Odds Ratio (OR):", input$or_other),
                              paste("Proportion in Controls:", input$p1_other),
                              paste("Case:Control Ratio:", input$r_other)
                            ),
                            "cohort" = c(
                              paste("Alpha (α):", input$alpha_other),
                              paste("Power (1-β):", input$power_other),
                              paste("Relative Risk (RR):", input$rr_other),
                              paste("Proportion in Unexposed:", input$p1_other),
                              paste("Exposed:Unexposed Ratio:", input$r_other)
                            ))
      
      steps <- c(steps, params_text,
                 paste("Non-response rate:", input$non_response_other, "%"),
                 "",
                 "CALCULATION DETAILS:",
                 output$formulaExplanation_other(),
                 "",
                 "ADJUSTMENT FOR NON-RESPONSE:",
                 output$adjustedSampleSize_other(),
                 "",
                 "STRATUM INFORMATION:"
      )
      
      for (i in 1:rv_other$stratumCount) {
        steps <- c(steps, 
                   paste("Stratum", i, ":", input[[paste0("stratum_other", i)]], 
                         "- Population:", input[[paste0("pop_other", i)]]))
      }
      
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
                 paste("Total samples:", sum(alloc$Proportional_Sample[1:rv_other$stratumCount])),
                 "",
                 "INTERPRETATION:",
                 interpretationText_other()
      )
      
      writeLines(steps, file)
    }
  )
  
  # Power Analysis section
  observeEvent(input$runPower, {
    tryCatch({
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
                       "One-Way ANOVA" = pwr.anova.test(k = 3, f = input$effectSize, sig.level = input$alpha, 
                                                        power = input$power),
                       "Two-Way ANOVA" = pwr.anova.test(k = 4, f = input$effectSize, sig.level = input$alpha, 
                                                        power = input$power),
                       "Proportion" = pwr.p.test(h = ES.h(input$effectSize, 0.5), sig.level = input$alpha, 
                                                 power = input$power, alternative = "two.sided"),
                       "Correlation" = pwr.r.test(r = input$effectSize, sig.level = input$alpha, 
                                                  power = input$power, alternative = "two.sided"),
                       "Chi-squared" = pwr.chisq.test(w = input$effectSize, df = 1, sig.level = input$alpha, 
                                                      power = input$power),
                       "Simple Linear Regression" = pwr.f2.test(u = 1, f2 = input$effectSize^2, 
                                                                sig.level = input$alpha, power = input$power),
                       "Multiple Linear Regression" = pwr.f2.test(u = input$predictors, f2 = input$effectSize^2, 
                                                                  sig.level = input$alpha, power = input$power)
      )
      
      output$powerResult <- renderText({
        paste(
          "Power Analysis Results for", input$testType, "\n",
          "========================================\n",
          "Effect size:", input$effectSize, "\n",
          "Alpha:", input$alpha, "\n",
          "Power:", input$power, "\n",
          "Required sample size:", ifelse(input$testType %in% c("Independent t-test", "Paired t-test"), 
                                          ceiling(result$n * 2), ceiling(result$n)), "\n",
          ifelse(input$testType %in% c("Independent t-test", "Paired t-test"), 
                 paste("Sample size per group:", ceiling(result$n)), "")
        )
      })
    }, error = function(e) {
      output$powerResult <- renderText({
        paste("Error in power analysis:", e$message)
      })
    })
  })
  
  output$downloadPowerSteps <- downloadHandler(
    filename = function() {
      paste("power_analysis_results_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      result_text <- capture.output({
        cat("POWER ANALYSIS RESULTS\n")
        cat("=====================\n")
        cat("Date:", Sys.Date(), "\n\n")
        cat("Test Type:", input$testType, "\n")
        cat("Effect Size:", input$effectSize, "\n")
        cat("Alpha:", input$alpha, "\n")
        cat("Power:", input$power, "\n")
        
        if (input$testType == "Multiple Linear Regression") {
          cat("Number of Predictors:", input$predictors, "\n")
        }
        
        cat("\n")
        
        # Run the analysis again to get results
        tryCatch({
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
                           "One-Way ANOVA" = pwr.anova.test(k = 3, f = input$effectSize, sig.level = input$alpha, 
                                                            power = input$power),
                           "Two-Way ANOVA" = pwr.anova.test(k = 4, f = input$effectSize, sig.level = input$alpha, 
                                                            power = input$power),
                           "Proportion" = pwr.p.test(h = ES.h(input$effectSize, 0.5), sig.level = input$alpha, 
                                                     power = input$power, alternative = "two.sided"),
                           "Correlation" = pwr.r.test(r = input$effectSize, sig.level = input$alpha, 
                                                      power = input$power, alternative = "two.sided"),
                           "Chi-squared" = pwr.chisq.test(w = input$effectSize, df = 1, sig.level = input$alpha, 
                                                          power = input$power),
                           "Simple Linear Regression" = pwr.f2.test(u = 1, f2 = input$effectSize^2, 
                                                                    sig.level = input$alpha, power = input$power),
                           "Multiple Linear Regression" = pwr.f2.test(u = input$predictors, f2 = input$effectSize^2, 
                                                                      sig.level = input$alpha, power = input$power)
          )
          
          cat("Required sample size:", ifelse(input$testType %in% c("Independent t-test", "Paired t-test"), 
                                              ceiling(result$n * 2), ceiling(result$n)), "\n")
          if (input$testType %in% c("Independent t-test", "Paired t-test")) {
            cat("Sample size per group:", ceiling(result$n), "\n")
          }
          cat("\nDetailed results:\n")
          print(result)
        }, error = function(e) {
          cat("Error in power analysis:", e$message, "\n")
        })
      })
      
      writeLines(result_text, file)
    }
  )
  
  # Descriptive Statistics section
  observeEvent(input$runDesc, {
    tryCatch({
      data_text <- input$dataInput
      if (nchar(trimws(data_text)) == 0) {
        output$unknownData <- renderText({
          "Please paste some data into the text area."
        })
        output$dataType <- reactive({"unknown"})
        return()
      }
      
      # Parse the data
      data_lines <- strsplit(data_text, "\n")[[1]]
      
      # Check if data has a header
      has_header <- length(data_lines) > 1 && 
        suppressWarnings(any(is.na(as.numeric(data_lines[1]))))
      
      if (has_header) {
        # Remove header for analysis
        data_values <- data_lines[-1]
      } else {
        data_values <- data_lines
      }
      
      # Clean data
      data_values <- trimws(data_values)
      data_values <- data_values[data_values != ""]
      
      if (length(data_values) == 0) {
        output$unknownData <- renderText({
          "No valid data found."
        })
        output$dataType <- reactive({"unknown"})
        return()
      }
      
      # Determine data type
      detect_data_type <- function(values) {
        # Try to convert to numeric
        numeric_test <- suppressWarnings(as.numeric(values))
        num_numeric <- sum(!is.na(numeric_test))
        prop_numeric <- num_numeric / length(values)
        
        # Check if values are Likert scale (ordinal)
        likert_pattern <- "^(strongly disagree|disagree|neutral|agree|strongly agree|[1-5])$"
        is_likert <- all(grepl(likert_pattern, tolower(values), ignore.case = TRUE))
        
        if (prop_numeric > 0.8) {
          return("numerical")
        } else if (is_likert) {
          return("ordinal")
        } else if (prop_numeric < 0.2 && length(unique(values)) < 10) {
          return("categorical")
        } else {
          return("mixed")
        }
      }
      
      data_type <- input$dataType
      if (data_type == "auto") {
        data_type <- detect_data_type(data_values)
      }
      
      output$dataType <- reactive({data_type})
      outputOptions(output, "dataType", suspendWhenHidden = FALSE)
      
      output$dataTypeDetection <- renderUI({
        div(
          class = "who-info-box",
          style = "margin-bottom: 20px;",
          HTML(paste0(
            "<strong>Detected Data Type:</strong> ", data_type, "<br>",
            "<strong>Number of observations:</strong> ", length(data_values), "<br>",
            "<strong>Missing values:</strong> ", sum(is.na(data_values) | data_values == ""), "<br>"
          ))
        )
      })
      
      # Generate appropriate summary based on data type
      if (data_type == "numerical") {
        numeric_values <- as.numeric(data_values)
        numeric_values <- numeric_values[!is.na(numeric_values)]
        
        if (length(numeric_values) == 0) {
          output$numericalSummary <- renderText({
            "No valid numerical data found."
          })
          return()
        }
        
        desc_stats <- list(
          "Number of observations" = length(numeric_values),
          "Number of missing values" = sum(is.na(numeric_values)),
          "Mean" = mean(numeric_values),
          "Median" = median(numeric_values),
          "Standard Deviation" = sd(numeric_values),
          "Variance" = var(numeric_values),
          "Minimum" = min(numeric_values),
          "Maximum" = max(numeric_values),
          "Range" = max(numeric_values) - min(numeric_values),
          "First Quartile (Q1)" = quantile(numeric_values, 0.25),
          "Third Quartile (Q3)" = quantile(numeric_values, 0.75),
          "Interquartile Range (IQR)" = IQR(numeric_values),
          "Skewness" = e1071::skewness(numeric_values),
          "Kurtosis" = e1071::kurtosis(numeric_values)
        )
        
        output$numericalSummary <- renderText({
          paste(
            "Numerical Data Summary\n",
            "======================\n",
            paste(names(desc_stats), sapply(desc_stats, function(x) round(x, 4)), sep = ": ", collapse = "\n"),
            "\n\nData preview (first 10 values):\n",
            paste(head(numeric_values, 10), collapse = ", ")
          )
        })
        
        output$numericalPlots <- renderPlot({
          par(mfrow = c(1, 2))
          hist(numeric_values, main = "Histogram", xlab = "Values", col = "#6BAED6")
          boxplot(numeric_values, main = "Boxplot", col = "#6BAED6")
        })
        
      } else if (data_type %in% c("categorical", "ordinal")) {
        # For categorical and ordinal data
        freq_table <- table(data_values)
        freq_percent <- prop.table(freq_table) * 100
        cum_freq <- cumsum(freq_table)
        
        # Create frequency table
        freq_df <- data.frame(
          Category = names(freq_table),
          Frequency = as.numeric(freq_table),
          Percentage = round(as.numeric(freq_percent), 2),
          Cumulative = as.numeric(cum_freq)
        )
        
        mode_val <- names(freq_table)[which.max(freq_table)]
        
        output$categoricalSummary <- renderText({
          paste(
            ifelse(data_type == "categorical", "Categorical", "Ordinal"), "Data Summary\n",
            "==================================\n",
            "Number of observations: ", length(data_values), "\n",
            "Number of categories: ", nrow(freq_df), "\n",
            "Mode: ", mode_val, " (", max(freq_table), " occurrences)\n\n",
            "Frequency Table:\n",
            paste(capture.output(print(freq_df, row.names = FALSE)), collapse = "\n")
          )
        })
        
        output$categoricalPlots <- renderPlot({
          par(mfrow = c(1, 2))
          barplot(freq_table, main = "Bar Plot", las = 2, col = "#6BAED6")
          pie(freq_table, main = "Pie Chart", col = rainbow(length(freq_table)))
        })
        
      } else if (data_type == "mixed") {
        output$mixedSummary <- renderText({
          "Mixed data types detected. Please specify the data type manually or clean your data."
        })
      }
      
    }, error = function(e) {
      output$unknownData <- renderText({
        paste("Error in data analysis:", e$message)
      })
      output$dataType <- reactive({"unknown"})
    })
  })
  
  output$downloadDescSteps <- downloadHandler(
    filename = function() {
      paste("descriptive_statistics_report_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      # Create progress indicator
      progress <- shiny::Progress$new()
      progress$set(message = "Generating Report", value = 0)
      on.exit(progress$close())
      
      tryCatch({
        # Create a new Word document
        doc <- officer::read_docx()
        
        # Add title
        doc <- doc %>% 
          officer::body_add_par("Descriptive Statistics Report", style = "heading 1") %>%
          officer::body_add_par(paste("Generated on:", Sys.Date()), style = "Normal") %>%
          officer::body_add_par("", style = "Normal")
        
        progress$set(value = 0.2, detail = "Processing data...")
        
        # Process the data (same logic as in runDesc)
        data_text <- input$dataInput
        if (nchar(trimws(data_text)) == 0) {
          doc <- doc %>% officer::body_add_par("No data available for analysis.", style = "Normal")
          print(doc, target = file)
          return()
        }
        
        data_lines <- strsplit(data_text, "\n")[[1]]
        has_header <- length(data_lines) > 1 && 
          suppressWarnings(any(is.na(as.numeric(data_lines[1]))))
        
        if (has_header) {
          data_values <- data_lines[-1]
        } else {
          data_values <- data_lines
        }
        
        data_values <- trimws(data_values)
        data_values <- data_values[data_values != ""]
        
        if (length(data_values) == 0) {
          doc <- doc %>% officer::body_add_par("No valid data found.", style = "Normal")
          print(doc, target = file)
          return()
        }
        
        # Determine data type
        detect_data_type <- function(values) {
          numeric_test <- suppressWarnings(as.numeric(values))
          num_numeric <- sum(!is.na(numeric_test))
          prop_numeric <- num_numeric / length(values)
          
          likert_pattern <- "^(strongly disagree|disagree|neutral|agree|strongly agree|[1-5])$"
          is_likert <- all(grepl(likert_pattern, tolower(values), ignore.case = TRUE))
          
          if (prop_numeric > 0.8) {
            return("numerical")
          } else if (is_likert) {
            return("ordinal")
          } else if (prop_numeric < 0.2 && length(unique(values)) < 10) {
            return("categorical")
          } else {
            return("mixed")
          }
        }
        
        data_type <- input$dataType
        if (data_type == "auto") {
          data_type <- detect_data_type(data_values)
        }
        
        # Add data type information
        doc <- doc %>%
          officer::body_add_par("Data Overview", style = "heading 2") %>%
          officer::body_add_par(paste("Data Type:", data_type), style = "Normal") %>%
          officer::body_add_par(paste("Number of observations:", length(data_values)), style = "Normal") %>%
          officer::body_add_par(paste("Missing values:", sum(is.na(data_values) | data_values == "")), style = "Normal") %>%
          officer::body_add_par("", style = "Normal")
        
        progress$set(value = 0.4, detail = "Generating statistics...")
        
        if (data_type == "numerical") {
          # Numerical data analysis
          numeric_values <- as.numeric(data_values)
          numeric_values <- numeric_values[!is.na(numeric_values)]
          
          if (length(numeric_values) > 0) {
            # Calculate statistics
            desc_stats <- data.frame(
              Statistic = c("Number of observations", "Mean", "Median", "Standard Deviation", 
                            "Variance", "Minimum", "Maximum", "Range", "First Quartile (Q1)", 
                            "Third Quartile (Q3)", "Interquartile Range (IQR)", "Skewness", "Kurtosis"),
              Value = c(length(numeric_values),
                        round(mean(numeric_values), 4),
                        round(median(numeric_values), 4),
                        round(sd(numeric_values), 4),
                        round(var(numeric_values), 4),
                        round(min(numeric_values), 4),
                        round(max(numeric_values), 4),
                        round(max(numeric_values) - min(numeric_values), 4),
                        round(quantile(numeric_values, 0.25), 4),
                        round(quantile(numeric_values, 0.75), 4),
                        round(IQR(numeric_values), 4),
                        round(e1071::skewness(numeric_values), 4),
                        round(e1071::kurtosis(numeric_values), 4))
            )
            
            # Add statistics table
            doc <- doc %>%
              officer::body_add_par("Descriptive Statistics", style = "heading 2")
            
            ft <- flextable::flextable(desc_stats) %>%
              flextable::theme_box() %>%
              flextable::autofit()
            doc <- flextable::body_add_flextable(doc, ft) %>%
              officer::body_add_par("", style = "Normal")
            
            progress$set(value = 0.6, detail = "Creating charts...")
            
            # Create and add charts
            temp_dir <- tempdir()
            
            # Histogram
            hist_file <- file.path(temp_dir, "histogram.png")
            png(hist_file, width = 6, height = 4, units = "in", res = 300)
            hist(numeric_values, main = "Histogram", xlab = "Values", col = "#6BAED6")
            dev.off()
            
            # Boxplot
            box_file <- file.path(temp_dir, "boxplot.png")
            png(box_file, width = 6, height = 4, units = "in", res = 300)
            boxplot(numeric_values, main = "Boxplot", col = "#6BAED6")
            dev.off()
            
            # Add charts to document
            doc <- doc %>%
              officer::body_add_par("Data Visualization", style = "heading 2") %>%
              officer::body_add_par("Histogram", style = "heading 3") %>%
              officer::body_add_img(hist_file, width = 6, height = 4) %>%
              officer::body_add_par("Boxplot", style = "heading 3") %>%
              officer::body_add_img(box_file, width = 6, height = 4)
            
            # Clean up temp files
            unlink(c(hist_file, box_file))
          }
          
        } else if (data_type %in% c("categorical", "ordinal")) {
          # Categorical/Ordinal data analysis
          freq_table <- table(data_values)
          freq_percent <- prop.table(freq_table) * 100
          cum_freq <- cumsum(freq_table)
          
          freq_df <- data.frame(
            Category = names(freq_table),
            Frequency = as.numeric(freq_table),
            Percentage = round(as.numeric(freq_percent), 2),
            Cumulative = as.numeric(cum_freq)
          )
          
          mode_val <- names(freq_table)[which.max(freq_table)]
          
          # Add frequency table
          doc <- doc %>%
            officer::body_add_par("Frequency Distribution", style = "heading 2") %>%
            officer::body_add_par(paste("Mode:", mode_val, "(", max(freq_table), "occurrences)"), style = "Normal")
          
          ft <- flextable::flextable(freq_df) %>%
            flextable::theme_box() %>%
            flextable::autofit()
          doc <- flextable::body_add_flextable(doc, ft) %>%
            officer::body_add_par("", style = "Normal")
          
          progress$set(value = 0.6, detail = "Creating charts...")
          
          # Create and add charts
          temp_dir <- tempdir()
          
          # Bar plot
          bar_file <- file.path(temp_dir, "barplot.png")
          png(bar_file, width = 8, height = 6, units = "in", res = 300)
          par(mar = c(7, 4, 4, 2) + 0.1)  # Increase bottom margin for long labels
          barplot(freq_table, main = "Bar Plot", las = 2, col = "#6BAED6")
          dev.off()
          
          # Pie chart
          pie_file <- file.path(temp_dir, "piechart.png")
          png(pie_file, width = 6, height = 6, units = "in", res = 300)
          pie(freq_table, main = "Pie Chart", col = rainbow(length(freq_table)))
          dev.off()
          
          # Add charts to document
          doc <- doc %>%
            officer::body_add_par("Data Visualization", style = "heading 2") %>%
            officer::body_add_par("Bar Plot", style = "heading 3") %>%
            officer::body_add_img(bar_file, width = 8, height = 6) %>%
            officer::body_add_par("Pie Chart", style = "heading 3") %>%
            officer::body_add_img(pie_file, width = 6, height = 6)
          
          # Clean up temp files
          unlink(c(bar_file, pie_file))
          
        } else {
          # Mixed or unknown data type
          doc <- doc %>%
            officer::body_add_par("Data Analysis", style = "heading 2") %>%
            officer::body_add_par("Mixed or unknown data type detected. Please specify the data type manually or clean your data.", style = "Normal") %>%
            officer::body_add_par("Raw data preview:", style = "heading 3") %>%
            officer::body_add_par(paste(head(data_values, 20), collapse = ", "), style = "Normal")
        }
        
        progress$set(value = 0.9, detail = "Finalizing document...")
        
        # Add data preview section
        doc <- doc %>%
          officer::body_add_par("Data Preview", style = "heading 2") %>%
          officer::body_add_par(paste("First 20 values:", paste(head(data_values, 20), collapse = ", ")), style = "Normal")
        
        # Save the document
        print(doc, target = file)
        
      }, error = function(e) {
        showNotification(paste("Error generating report:", e$message), type = "error")
      })
    }
  )
}



# Add JavaScript for localStorage functionality
jscode <- "
shinyjs.saveValues = function(params) {
  localStorage.setItem(params.section, params.values);
}

shinyjs.clearValues = function(params) {
  localStorage.removeItem(params);
}

// Initialize values from localStorage when page loads
$(document).on('shiny:connected', function(event) {
  // Proportional Allocation
  var propValues = localStorage.getItem('proportional');
  if (propValues) {
    Shiny.setInputValue('proportional_values', propValues);
  }
  
  // Taro Yamane
  var yamaneValues = localStorage.getItem('yamane');
  if (yamaneValues) {
    Shiny.setInputValue('yamane_values', yamaneValues);
  }
  
  // Cochran
  var cochranValues = localStorage.getItem('cochran');
  if (cochranValues) {
    Shiny.setInputValue('cochran_values', cochranValues);
  }
  
  // Other Formulas
  var otherValues = localStorage.getItem('other');
  if (otherValues) {
    Shiny.setInputValue('other_values', otherValues);
  }
  
  // Power Analysis
  var powerValues = localStorage.getItem('power');
  if (powerValues) {
    Shiny.setInputValue('power_values', powerValues);
  }
  
  // Descriptive Statistics
  var descValues = localStorage.getItem('desc');
  if (descValues) {
    Shiny.setInputValue('desc_values', descValues);
  }
});
"

# Run the application
shinyApp(ui = ui, server = server)
