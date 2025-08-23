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

ui <- navbarPage(
  title = div(icon("chart-bar"), "CalcuStats"),
  theme = shinytheme("cosmo"),
  header = div(
    style = "text-align: right; padding: 5px 15px; background-color: #f8f9fa; font-size: 14px;",
    textOutput("welcomeMessage"),
    textOutput("currentDateTime")
  ),
  footer = div(
    style = "text-align: center; padding: 10px; background-color: #f8f9fa; border-top: 1px solid #ddd; font-size: 14px;",
    HTML("<p>© 2025 Mudasir Mohammed Ibrahim. All rights reserved. | 
         <a href='https://github.com/mudassiribrahim30' target='_blank'>GitHub Profile</a></p>"),
    div(
      style = "margin-top: 10px; font-size: 0.9em; color: #555;",
      "Your Companion for Sample Size Calculation and Descriptive Analytics"
    )
  ),
  useShinyjs(),
  
  tabPanel("Proportional Allocation",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "proportional_sidebar",
               div(
                 style = "margin-bottom: 15px;",
                 actionButton("reset_proportional", "Reset All Values", class = "btn-danger",
                              icon = icon("refresh"))
               ),
               numericInput("custom_sample", "Your Sample Size (n)", value = 100, min = 1, step = 1),
               numericInput("non_response_custom", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
               actionButton("addStratum_custom", "Add Stratum", class = "btn-primary"),
               actionButton("removeStratum_custom", "Remove Last Stratum", class = "btn-warning"),
               uiOutput("stratumInputs_custom"),
               div(
                 style = "margin-top: 10px; padding: 10px; background-color: #f8f9fa; border-radius: 5px; font-size: 14px;",
                 p("For a single population, enter the total size below."),
                 p("For multiple populations, click 'Add Stratum' to create additional groups.")
               ),
               conditionalPanel(
                 condition = "output.stratumCount_custom > 1",
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
                     
                     actionButton("updateFlowchart_custom", "Update Diagram", class = "btn-primary"),
                     br(), br(),
                     h4("Diagram Editor"),
                     div(
                       style = "border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;",
                       h5("Text Editing"),
                       textInput("nodeText_custom", "Current Text:", ""),
                       textInput("newText_custom", "New Text (use -> for replacement):", ""),
                       actionButton("editNodeText_custom", "Update Text", class = "btn-info"),
                       actionButton("deleteText_custom", "Delete Text", class = "btn-danger"),
                       actionButton("resetDiagramText_custom", "Reset All Text", class = "btn-warning")
                     ),
                     div(
                       style = "border: 1px solid #ddd; padding: 10px;",
                       h5("Font Size Controls"),
                       sliderInput("selectedNodeFontSize_custom", "Selected Node Font Size", 
                                   min = 12, max = 24, value = 16),
                       sliderInput("selectedEdgeFontSize_custom", "Selected Edge Font Size", 
                                   min = 10, max = 20, value = 14),
                       actionButton("applyFontSizes_custom", "Apply Font Sizes", class = "btn-primary")
                     ),
                     br(), br(),
                     h4("Download Diagram"),
                     div(style = "display: inline-block;", 
                         downloadButton("downloadFlowchartPNG_custom", "PNG (High Quality)", class = "btn-success")),
                     div(style = "display: inline-block; margin-left: 5px;", 
                         downloadButton("downloadFlowchartSVG_custom", "SVG (Vector)", class = "btn-success")),
                     div(style = "display: inline-block; margin-left: 5px;", 
                         downloadButton("downloadFlowchartPDF_custom", "PDF (Vector)", class = "btn-success"))
                   )
                 )
               ),
               downloadButton("downloadCustomWord", "Download as Word", class = "btn-primary"),
               downloadButton("downloadCustomSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Proportional Allocation for Specified Sample Size"),
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
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
               tableOutput("allocationTableCustom"),
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
                 style = "margin-bottom: 15px;",
                 actionButton("reset_yamane", "Reset All Values", class = "btn-danger",
                              icon = icon("refresh"))
               ),
               numericInput("e", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
               numericInput("non_response", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
               actionButton("addStratum", "Add Stratum", class = "btn-primary"),
               actionButton("removeStratum", "Remove Last Stratum", class = "btn-warning"),
               uiOutput("stratumInputs"),
               div(
                 style = "margin-top: 10px; padding: 10px; background-color: #f8f9fa; border-radius: 5px; font-size: 14px;",
                 p("For a single population, enter the total size below."),
                 p("For multiple populations, click 'Add Stratum' to create additional groups.")
               ),
               conditionalPanel(
                 condition = "output.stratumCount > 1",
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
                     
                     actionButton("updateFlowchart", "Update Diagram", class = "btn-primary"),
                     br(), br(),
                     h4("Diagram Editor"),
                     div(
                       style = "border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;",
                       h5("Text Editing"),
                       textInput("nodeText", "Current Text:", ""),
                       textInput("newText", "New Text (use -> for replacement):", ""),
                       actionButton("editNodeText", "Update Text", class = "btn-info"),
                       actionButton("deleteText", "Delete Text", class = "btn-danger"),
                       actionButton("resetDiagramText", "Reset All Text", class = "btn-warning")
                     ),
                     div(
                       style = "border: 1px solid #ddd; padding: 10px;",
                       h5("Font Size Controls"),
                       sliderInput("selectedNodeFontSize", "Selected Node Font Size", 
                                   min = 12, max = 24, value = 16),
                       sliderInput("selectedEdgeFontSize", "Selected Edge Font Size", 
                                   min = 10, max = 20, value = 14),
                       actionButton("applyFontSizes", "Apply Font Sizes", class = "btn-primary")
                     ),
                     br(), br(),
                     h4("Download Diagram"),
                     div(style = "display: inline-block;", 
                         downloadButton("downloadFlowchartPNG", "PNG (High Quality)", class = "btn-success")),
                     div(style = "display: inline-block; margin-left: 5px;", 
                         downloadButton("downloadFlowchartSVG", "SVG (Vector)", class = "btn-success")),
                     div(style = "display: inline-block; margin-left: 5px;", 
                         downloadButton("downloadFlowchartPDF", "PDF (Vector)", class = "btn-success"))
                   )
                 )
               ),
               downloadButton("downloadWord", "Download as Word", class = "btn-primary"),
               downloadButton("downloadSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Taro Yamane Method with Proportional Allocation"),
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
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
               tableOutput("allocationTable"),
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
                 style = "margin-bottom: 15px;",
                 actionButton("reset_cochran", "Reset All Values", class = "btn-danger",
                              icon = icon("refresh"))
               ),
               numericInput("p", "Estimated Proportion (p)", value = 0.5, min = 0.01, max = 0.99, step = 0.01),
               numericInput("z", "Z-score (Z)", value = 1.96),
               numericInput("e_c", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
               numericInput("N_c", "Population Size (optional)", value = NULL, min = 1),
               numericInput("non_response_c", "Non-response Rate (%)", value = 0, min = 0, max = 100, step = 1),
               actionButton("addStratum_c", "Add Stratum", class = "btn-primary"),
               actionButton("removeStratum_c", "Remove Last Stratum", class = "btn-warning"),
               uiOutput("stratumInputs_c"),
               conditionalPanel(
                 condition = "output.stratumCount_c > 1",
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
                     
                     actionButton("updateFlowchart_c", "Update Diagram", class = "btn-primary"),
                     br(), br(),
                     h4("Diagram Editor"),
                     div(
                       style = "border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;",
                       h5("Text Editing"),
                       textInput("nodeText_c", "Current Text:", ""),
                       textInput("newText_c", "New Text (use -> for replacement):", ""),
                       actionButton("editNodeText_c", "Update Text", class = "btn-info"),
                       actionButton("deleteText_c", "Delete Text", class = "btn-danger"),
                       actionButton("resetDiagramText_c", "Reset All Text", class = "btn-warning")
                     ),
                     div(
                       style = "border: 1px solid #ddd; padding: 10px;",
                       h5("Font Size Controls"),
                       sliderInput("selectedNodeFontSize_c", "Selected Node Font Size", 
                                   min = 12, max = 24, value = 16),
                       sliderInput("selectedEdgeFontSize_c", "Selected Edge Font Size", 
                                   min = 10, max = 20, value = 14),
                       actionButton("applyFontSizes_c", "Apply Font Sizes", class = "btn-primary")
                     ),
                     br(), br(),
                     h4("Download Diagram"),
                     div(style = "display: inline-block;", 
                         downloadButton("downloadFlowchartPNG_c", "PNG (High Quality)", class = "btn-success")),
                     div(style = "display: inline-block; margin-left: 5px;", 
                         downloadButton("downloadFlowchartSVG_c", "SVG (Vector)", class = "btn-success")),
                     div(style = "display: inline-block; margin-left: 5px;", 
                         downloadButton("downloadFlowchartPDF_c", "PDF (Vector)", class = "btn-success"))
                   )
                 )
               ),
               downloadButton("downloadCochranWord", "Download as Word", class = "btn-primary"),
               downloadButton("downloadCochranSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Cochran Sample Size Calculator with Proportional Allocation"),
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px; font-size: 14px;",
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
               tableOutput("cochranAllocationTable"),
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
  
  tabPanel("Power Analysis",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               id = "power_sidebar",
               div(
                 style = "margin-bottom: 15px;",
                 actionButton("reset_power", "Reset All Values", class = "btn-danger",
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
               actionButton("runPower", "Run Power Analysis", class = "btn-primary"),
               downloadButton("downloadPowerSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Power Analysis for Inferential Tests"),
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
                 style = "margin-bottom: 15px;",
                 actionButton("reset_desc", "Reset All Values", class = "btn-danger",
                              icon = icon("refresh"))
               ),
               tags$textarea(id = "dataInput", rows = 10, cols = 30,
                             placeholder = "Paste a column of data (with header) from Excel or statistical software...",
                             style = "font-size: 14px;"),
               actionButton("runDesc", "Get Descriptive Statistics", class = "btn-primary"),
               downloadButton("downloadDescSteps", "Download Calculation Steps", class = "btn-info")
             ),
             mainPanel(
               width = 8,
               h3("Descriptive Statistics"),
               div(style = "font-size: 14px;", verbatimTextOutput("descResult"))
             )
           )
  ),
  
  tabPanel("How to Use App",
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
                 tags$li(strong("Power Analysis:"), "Determine required sample size for various statistical tests."),
                 tags$li(strong("Descriptive Statistics:"), "Compute basic statistics for your data.")
               ),
               
               h3("About the Developer"),
               p("CalcuStats was developed by Mudasir Mohammed Ibrahim, a Registered Nurse with a Bachelor of Science degree and Diploma qualifications."),
               p("With expertise in both healthcare and data analysis, Mudasir created this tool to help researchers and students perform essential statistical calculations with ease."),
               p("Connect with Mudasir on GitHub:", 
                 tags$a(href="https://github.com/mudassiribrahim30", target="_blank", "github.com/mudassiribrahim30")),
               
               h3("Detailed Instructions"),
               h4("Proportional Allocation"),
               tags$ol(
                 tags$li("Enter your desired sample size in the 'Your Sample Size (n)' field."),
                 tags$li("Adjust the non-response rate if needed."),
                 tags$li("Add strata (population groups) using the 'Add Stratum' button."),
                 tags$li("For each stratum, provide a name and population size."),
                 tags$li("The app will automatically calculate the proportional allocation."),
                 tags$li("For multiple strata, you can generate a flow chart diagram."),
                 tags$li("Download results as a Word document or calculation steps.")
               ),
               
               h4("Taro Yamane Method"),
               tags$ol(
                 tags$li("Enter your margin of error (e.g., 0.05 for 5%)."),
                 tags$li("Adjust the non-response rate if needed."),
                 tags$li("Add strata as needed and provide population sizes."),
                 tags$li("The app calculates the required sample size using the formula: n = N / (1 + N*e²)"),
                 tags$li("Results show the proportional allocation across strata."),
                 tags$li("Flow chart visualization is available for multiple strata.")
               ),
               
               h4("Cochran Formula"),
               tags$ol(
                 tags$li("Enter the estimated proportion (p) - use 0.5 for maximum variability."),
                 tags$li("Set the Z-score (1.96 for 95% confidence)."),
                 tags$li("Enter your desired margin of error."),
                 tags$li("Optionally provide the population size for finite correction."),
                 tags$li("Add strata if needed for proportional allocation."),
                 tags$li("The app calculates the sample size using Cochran's formula.")
               ),
               
               h4("Power Analysis"),
               tags$ol(
                 tags$li("Select your statistical test type."),
                 tags$li("Enter the effect size (Cohen's d for t-tests, f for ANOVA, etc.)."),
                 tags$li("Set your significance level (alpha, typically 0.05)."),
                 tags$li("Enter your desired power (typically 0.8 or 0.9)."),
                 tags$li("Click 'Run Power Analysis' to see results."),
                 tags$li("For regression, specify number of predictors if needed.")
               ),
               
               h4("Descriptive Statistics"),
               tags$ol(
                 tags$li("Paste your data (one column with header) into the text area."),
                 tags$li("Click 'Get Descriptive Statistics'."),
                 tags$li("The app will compute mean, SD, min, max, etc. for numeric data."),
                 tags$li("For categorical data, it provides frequency tables.")
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
  
  tabPanel("Usage",
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
    if (n == 0) return(0)
    non_response_rate <- input$non_response_custom / 100
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
    if (is.null(alloc)) return("")
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
    filename = function() paste("Proportional_Allocation_Report_", Sys.Date(), ".docx", sep = ""),
    content = function(file) {
      doc <- read_docx() %>%
        body_add_par("CalcuStats", style = "heading 1") %>%
        body_add_par("Proportional Allocation for Specified Sample Size", style = "heading 2") %>%
        body_add_par(paste("Total Population:", totalPopulation_custom()), style = "Normal") %>%
        body_add_par(paste("Your Specified Sample Size:", input$custom_sample), style = "Normal") %>%
        body_add_par(paste("Non-response Rate:", input$non_response_custom, "%"), style = "Normal") %>%
        body_add_par(paste("Adjusted Sample Size (accounting for non-response):", adjustedSampleSize_custom()), style = "Normal") %>%
        body_add_par("Proportional Allocation Table:", style = "heading 2") %>%
        body_add_par("The proportional allocation formula is: n_i = (N_i / N) × n", style = "Normal")
      ft <- flextable(allocationData_custom()) %>% autofit()
      doc <- doc %>% body_add_flextable(ft)
      doc <- doc %>%
        body_add_par("Interpretation", style = "heading 2") %>%
        body_add_par(interpretationText_custom(), style = "Normal")
      print(doc, target = file)
    }
  )
  
  output$downloadCustomSteps <- downloadHandler(
    filename = function() paste("Proportional_Allocation_Steps_", Sys.Date(), ".txt", sep = ""),
    content = function(file) {
      steps <- c(
        "Proportional Allocation Calculation Steps",
        "=======================================",
        "",
        paste("Total Population (N):", totalPopulation_custom()),
        paste("Your Specified Sample Size (n):", input$custom_sample),
        paste("Non-response Rate:", input$non_response_custom, "%"),
        "",
        "1. Non-response Adjustment:",
        paste("Adjusted sample size = ", input$custom_sample, " / (1 - ", input$non_response_custom/100, ")"),
        paste("Calculation: ", input$custom_sample, " / ", (1 - input$non_response_custom/100), " = ", input$custom_sample / (1 - input$non_response_custom/100)),
        paste("Final adjusted sample size: ", adjustedSampleSize_custom()),
        "",
        "2. Proportional Allocation:",
        capture.output(print(allocationData_custom())),
        "",
        "Interpretation:",
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
  
  sampleSize <- reactive({
    N <- totalPopulation()
    e <- input$e
    if (N == 0 || e == 0) return(0)
    ceiling(N / (1 + N * e^2))
  })
  
  adjustedSampleSize <- reactive({
    n <- sampleSize()
    if (n == 0) return(0)
    non_response_rate <- input$non_response / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  output$formulaExplanation <- renderText({
    N <- totalPopulation()
    e <- input$e
    if (N == 0 || e == 0) return("")
    
    denominator <- 1 + N * e^2
    raw_sample <- N / denominator
    
    paste(
      "Yamane Formula Calculation Steps:\n",
      "1. Formula: n = N / (1 + N*e²)\n",
      "2. Total Population (N) = ", N, "\n",
      "3. Margin of Error (e) = ", e, "\n",
      "4. Calculate denominator: 1 + N*e² = 1 + (", N, " * ", e, "²) = 1 + (", N, " * ", e^2, ") = ", denominator, "\n",
      "5. Calculate raw sample size: ", N, " / ", denominator, " = ", raw_sample, "\n",
      "6. Apply ceiling function to round up to nearest integer: ", ceiling(raw_sample), "\n",
      "\nFinal calculated sample size after rounding: ", sampleSize()
    )
  })
  
  output$adjustedSampleSize <- renderText({
    n <- sampleSize()
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
    paste0("Based on a total population of ", totalPopulation(),
           ", a margin of error of ", input$e,
           ", and a non-response rate of ", input$non_response, "%",
           ", the required sample size is ", adjustedSampleSize(), ". ",
           paste(sentences, collapse = " "))
  })
  
  output$totalPop <- renderText({ paste("Total Population:", totalPopulation()) })
  output$sampleSize <- renderText({ paste("Initial Sample Size:", sampleSize()) })
  output$allocationTable <- renderTable({ allocationData() })
  output$interpretationText <- renderText({ interpretationText() })
  
  generateDotCode_yamane <- function() {
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
    
    dot_code <- generateDotCode_yamane()
    grViz(dot_code)
  })
  
  observeEvent(input$resetDiagramText, {
    rv$nodeTexts <- list()
    rv$edgeTexts <- list()
    rv$nodeFontSizes <- list()
    rv$edgeFontSizes <- list()
    updateTextInput(session, "nodeText", value = "")
    updateTextInput(session, "newText", value = "")
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode_yamane()
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
      dot_code <- generateDotCode_yamane()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetEdgeText, {
    rv$edgeTexts[[rv$selectedEdge]] <- NULL
    rv$edgeFontSizes[[rv$selectedEdge]] <- NULL
    updateTextInput(session, "nodeText", value = "")
    output$flowchart <- renderGrViz({
      dot_code <- generateDotCode_yamane()
      grViz(dot_code)
    })
  })
  
  output$downloadFlowchartPNG <- downloadHandler(
    filename = function() {
      paste("taro_yamane_flowchart_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_yamane()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_png(tmp_file, file, width = 3000, height = 2000)
      unlink(tmp_file)
    }
  )
  
  output$downloadFlowchartSVG <- downloadHandler(
    filename = function() {
      paste("taro_yamane_flowchart_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_yamane()
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(file)
    }
  )
  
  output$downloadFlowchartPDF <- downloadHandler(
    filename = function() {
      paste("taro_yamane_flowchart_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_yamane()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_pdf(tmp_file, file)
      unlink(tmp_file)
    }
  )
  
  output$downloadWord <- downloadHandler(
    filename = function() paste("Sample_Size_Report_", Sys.Date(), ".docx", sep = ""),
    content = function(file) {
      doc <- read_docx() %>%
        body_add_par("CalcuStats", style = "heading 1") %>%
        body_add_par("Sample Size Calculation (Taro Yamane Method)", style = "heading 2") %>%
        body_add_par(paste("Total Population:", totalPopulation()), style = "Normal") %>%
        body_add_par(paste("Margin of Error:", input$e), style = "Normal") %>%
        body_add_par(paste("Non-response Rate:", input$non_response, "%"), style = "Normal") %>%
        body_add_par(paste("Initial Sample Size:", sampleSize()), style = "Normal") %>%
        body_add_par(paste("Adjusted Sample Size (accounting for non-response):", adjustedSampleSize()), style = "Normal") %>%
        body_add_par("Proportional Allocation Table:", style = "heading 2") %>%
        body_add_par("The proportional allocation formula is: n_i = (N_i / N) × n", style = "Normal")
      ft <- flextable(allocationData()) %>% autofit()
      doc <- doc %>% body_add_flextable(ft)
      doc <- doc %>%
        body_add_par("Interpretation", style = "heading 2") %>%
        body_add_par(interpretationText(), style = "Normal")
      print(doc, target = file)
    }
  )
  
  output$downloadSteps <- downloadHandler(
    filename = function() {
      paste("Taro_Yamane_Calculation_Steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      steps <- c(
        "Taro Yamane Sample Size Calculation Steps",
        "========================================",
        "",
        paste("Total Population (N):", totalPopulation()),
        paste("Margin of Error (e):", input$e),
        paste("Non-response Rate:", input$non_response, "%"),
        "",
        "1. Basic Yamane Formula Calculation:",
        "Formula: n = N / (1 + N*e²)",
        paste("Calculation: ", totalPopulation(), " / (1 + ", totalPopulation(), "*", input$e, "^2)"),
        paste("Denominator: 1 + ", totalPopulation(), "*", input$e^2, " = ", 1 + totalPopulation() * input$e^2),
        paste("Raw sample size: ", totalPopulation(), " / ", (1 + totalPopulation() * input$e^2), 
              " = ", round(totalPopulation() / (1 + totalPopulation() * input$e^2), 2)),
        paste("Rounded sample size: ", sampleSize()),
        "",
        "2. Non-response Adjustment:",
        paste("Adjusted sample size = ", sampleSize(), " / (1 - ", input$non_response / 100, ")"),
        paste("Calculation: ", sampleSize(), " / ", (1 - input$non_response / 100), 
              " = ", round(sampleSize() / (1 - input$non_response / 100), 2)),
        paste("Final adjusted sample size: ", adjustedSampleSize()),
        "",
        "3. Proportional Allocation:"
      )
      
      # Append the output from allocationData() as character lines
      allocation_lines <- capture.output(print(allocationData()))
      full_output <- c(steps, allocation_lines)
      
      # Write to file
      writeLines(full_output, file)
    }
  )
  
  
  # Cochran section variables
  rv_c <- reactiveValues(
    stratumCount = 1,
    selectedNode = NULL,
    selectedEdge = NULL,
    nodeTexts = list(),
    edgeTexts = list()
  )
  
  observeEvent(input$addStratum_c, { rv_c$stratumCount <- rv_c$stratumCount + 1 })
  observeEvent(input$removeStratum_c, { if (rv_c$stratumCount > 1) rv_c$stratumCount <- rv_c$stratumCount - 1 })
  
  output$stratumInputs_c <- renderUI({
    lapply(1:rv_c$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum_c", i),
                      label = paste("Stratum", i, "Name"),
                      value = ifelse(i <= length(initValues$cochran$stratum_names), 
                                     initValues$cochran$stratum_names[i], 
                                     paste("Stratum", i)),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop_c", i),
                         label = paste("Stratum", i, "Population"),
                         value = ifelse(i <= length(initValues$cochran$stratum_pops), 
                                        initValues$cochran$stratum_pops[i], 
                                        100),
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  output$stratumCount_c <- reactive({
    rv_c$stratumCount
  })
  outputOptions(output, "stratumCount_c", suspendWhenHidden = FALSE)
  
  totalPopulation_c <- reactive({
    sum(sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]]), na.rm = TRUE)
  })
  
  cochranSampleSize <- reactive({
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- if (is.null(input$N_c) || is.na(input$N_c)) NULL else input$N_c
    
    n0 <- (z^2 * p * (1 - p)) / (e^2)
    
    if (!is.null(N) && N > 0) {
      n <- n0 / (1 + (n0 - 1)/N)
      return(list(initial = n0, corrected = n, finite_correction_applied = TRUE))
    } else {
      return(list(initial = n0, corrected = n0, finite_correction_applied = FALSE))
    }
  })
  
  adjustedCochranSampleSize <- reactive({
    res <- cochranSampleSize()
    n <- res$corrected
    if (n == 0) return(0)
    non_response_rate <- input$non_response_c / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  cochranAllocationData <- reactive({
    N <- if (is.null(input$N_c) || is.na(input$N_c)) totalPopulation_c() else input$N_c
    n <- adjustedCochranSampleSize()
    strata <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]] )
    pops <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]] )
    
    if (length(pops) == 0 || sum(pops) == 0) return(NULL)
    
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
  
  cochranInterpretationText <- reactive({
    res <- cochranSampleSize()
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- if (is.null(input$N_c) || is.na(input$N_c)) totalPopulation_c() else input$N_c
    alloc <- cochranAllocationData()
    
    if (is.null(alloc)) {
      interpretation <- paste(
        "Assuming a population proportion of", p,
        ", a confidence level corresponding to Z =", z,
        " (", round(pnorm(z) * 200 - 100, 1), "% confidence),",
        " and a margin of error of", e,
        ", the required sample size is", ceiling(res$corrected), "."
      )
      
      if (res$finite_correction_applied) {
        interpretation <- paste0(interpretation,
                                 " This includes a finite population correction for N = ", N, "."
        )
      } else {
        interpretation <- paste0(interpretation,
                                 " No finite population correction was applied (assumes infinite population)."
        )
      }
      
      return(interpretation)
    }
    
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    
    interpretation <- paste(
      "Assuming a population proportion of", p,
      ", a confidence level corresponding to Z =", z,
      " (", round(pnorm(z) * 200 - 100, 1), "% confidence),",
      " and a margin of error of", e,
      ", the required sample size is", adjustedCochranSampleSize(), "."
    )
    
    if (res$finite_correction_applied) {
      interpretation <- paste0(interpretation,
                               " This includes a finite population correction for N = ", N, "."
      )
    } else {
      interpretation <- paste0(interpretation,
                               " No finite population correction was applied (assumes infinite population)."
      )
    }
    
    paste0(interpretation, " ", paste(sentences, collapse = " "))
  })
  
  output$cochranFormulaExplanation <- renderText({
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- if (is.null(input$N_c) || is.na(input$N_c)) NULL else input$N_c
    res <- cochranSampleSize()
    
    calculation_steps <- paste(
      "Cochran's Formula Calculation Steps:\n",
      "1. Basic formula: n₀ = (Z² × p × (1-p)) / e²\n",
      "2. Z-score (Z) = ", z, "\n",
      "3. Estimated proportion (p) = ", p, "\n",
      "4. Margin of error (e) = ", e, "\n",
      "5. Calculate numerator (Z² × p × (1-p)): (", z, "² × ", p, " × ", (1-p), ") = ", 
      (z^2 * p * (1 - p)), "\n",
      "6. Calculate denominator (e²): ", e, "² = ", e^2, "\n",
      "7. Initial sample size (n₀): ", (z^2 * p * (1 - p)), " / ", e^2, " = ", res$initial, "\n"
    )
    
    if (res$finite_correction_applied) {
      calculation_steps <- paste0(calculation_steps,
                                  "\nFinite Population Correction Applied (N = ", N, "):\n",
                                  "8. Correction formula: n = n₀ / (1 + (n₀ - 1)/N)\n",
                                  "9. Calculate denominator: 1 + (", res$initial, " - 1)/", N, " = ", 
                                  (1 + (res$initial - 1)/N), "\n",
                                  "10. Corrected sample size: ", res$initial, " / ", (1 + (res$initial - 1)/N), " = ", res$corrected, "\n"
      )
    } else {
      calculation_steps <- paste0(calculation_steps,
                                  "\nNo finite population correction applied (population size not specified or infinite).\n"
      )
    }
    
    calculation_steps
  })
  
  output$cochranAdjustedSample <- renderText({
    res <- cochranSampleSize()
    adj_n <- adjustedCochranSampleSize()
    non_response_rate <- input$non_response_c
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", ceiling(res$corrected), "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", ceiling(res$corrected), " / (1 - ", non_response_rate/100, ") = ", ceiling(res$corrected) / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  output$cochranSample <- renderText({
    res <- cochranSampleSize()
    if (res$finite_correction_applied) {
      paste("Initial Sample Size (n₀):", ceiling(res$initial), "\n",
            "Corrected Sample Size (n):", ceiling(res$corrected))
    } else {
      paste("Sample Size (n₀):", ceiling(res$initial))
    }
  })
  
  output$cochranAllocationTable <- renderTable({
    cochranAllocationData()
  })
  
  output$cochranInterpretation <- renderText({
    cochranInterpretationText()
  })
  
  generateDotCode_cochran <- function() {
    if (rv_c$stratumCount <= 1) return("")
    
    strata <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]])
    pops <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]])
    alloc <- cochranAllocationData()
    total_sample <- sum(alloc$Proportional_Sample[1:rv_c$stratumCount])
    
    nodes <- paste0("node", 1:(rv_c$stratumCount + 3))
    
    # Create node labels based on user preferences
    node_labels <- c(paste0("Total Population\nN = ", ifelse(is.null(input$N_c) || is.na(input$N_c), totalPopulation_c(), input$N_c)))
    
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
    
    dot_code <- paste0(
      "digraph flowchart {
        rankdir=TB;
        layout=\"", input$flowchartLayout_c, "\";
        node [fontname=Arial, shape=\"", input$nodeShape_c, "\", style=filled, fillcolor='", input$nodeColor_c, "', 
              width=", input$nodeWidth_c, ", height=", input$nodeHeight_c, ", fontsize=", input$nodeFontSize_c, "];
        edge [color='", input$edgeColor_c, "', fontsize=", input$edgeFontSize_c, ", arrowsize=", input$arrowSize_c, "];
        
        // Nodes
        '", nodes[1], "' [label='", node_labels[1], "', id='", nodes[1], "'];
        '", nodes[length(nodes)], "' [label='", node_labels[length(node_labels)], "', id='", nodes[length(nodes)], "'];
      ",
      paste0(sapply(1:rv_c$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' [label='", node_labels[i+1], "', id='", nodes[i+1], "'];")
      }), collapse = "\n"),
      "
        
        // Edges
        '", nodes[1], "' -> {",
      paste0("'", nodes[2:(rv_c$stratumCount+1)], "'", collapse = " "),
      "} [label='', id='stratification_edges'];
      ",
      paste0(sapply(1:rv_c$stratumCount, function(i) {
        paste0("'", nodes[i+1], "' -> '", nodes[length(nodes)], "' [label='", edge_labels[i], "', id='edge", i, "'];")
      }), collapse = "\n"),
      "
      }"
    )
    
    return(dot_code)
  }
  
  output$flowchart_c <- renderGrViz({
    if (!input$showFlowchart_c || rv_c$stratumCount <= 1) return()
    
    dot_code <- generateDotCode_cochran()
    grViz(dot_code)
  })
  
  jsCode_cochran <- '
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
    session$sendCustomMessage(type='jsCode', list(value = jsCode_cochran))
  })
  
  observeEvent(input$selected_node_c, {
    rv_c$selectedNode <- input$selected_node_c
    rv_c$selectedEdge <- NULL
    
    if (!is.null(rv_c$nodeTexts[[rv_c$selectedNode]])) {
      updateTextInput(session, "nodeText_c", value = rv_c$nodeTexts[[rv_c$selectedNode]])
    } else {
      node_index <- as.numeric(gsub("node", "", rv_c$selectedNode))
      if (node_index == 1) {
        updateTextInput(session, "nodeText_c", 
                        value = paste0("Total Population\nN = ", 
                                       ifelse(is.null(input$N_c) || is.na(input$N_c), totalPopulation_c(), input$N_c)))
      } else if (node_index == rv_c$stratumCount + 2) {
        alloc <- cochranAllocationData()
        total_sample <- sum(alloc$Proportional_Sample[1:rv_c$stratumCount])
        updateTextInput(session, "nodeText_c", 
                        value = paste0("Total Sample\nn = ", total_sample))
      } else {
        stratum_num <- node_index - 1
        strata <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("stratum_c", i)]])
        pops   <- sapply(1:rv_c$stratumCount, function(i) input[[paste0("pop_c", i)]])
        alloc  <- cochranAllocationData()
        
        updateTextInput(session, "nodeText_c", 
                        value = paste0("Stratum: ", strata[stratum_num], 
                                       "\nPop: ", pops[stratum_num], 
                                       "\nSample: ", alloc$Proportional_Sample[stratum_num]))
      }
    }
  })
  
  
  observeEvent(input$selected_edge_c, {
    rv_c$selectedEdge <- input$selected_edge_c
    rv_c$selectedNode <- NULL
    
    if (!is.null(rv_c$edgeTexts[[rv_c$selectedEdge]])) {
      updateTextInput(session, "nodeText_c", value = rv_c$edgeTexts[[rv_c$selectedEdge]])
    } else {
      edge_num <- as.numeric(gsub("edge", "", rv_c$selectedEdge))
      alloc <- cochranAllocationData()
      updateTextInput(session, "nodeText_c", 
                      value = paste0("", alloc$Proportional_Sample[edge_num]))
    }
  })
  
  observeEvent(input$editNodeText_c, {
    if (!is.null(rv_c$selectedNode)) {
      rv_c$nodeTexts[[rv_c$selectedNode]] <- input$nodeText_c
    } else if (!is.null(rv_c$selectedEdge)) {
      rv_c$edgeTexts[[rv_c$selectedEdge]] <- input$nodeText_c
    }
    
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_cochran()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetDiagramText_c, {
    rv_c$nodeTexts <- list()
    rv_c$edgeTexts <- list()
    updateTextInput(session, "nodeText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_cochran()
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
    updateTextInput(session, "nodeText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_cochran()
      grViz(dot_code)
    })
  })
  
  observeEvent(input$resetEdgeText_c, {
    rv_c$edgeTexts[[rv_c$selectedEdge]] <- NULL
    updateTextInput(session, "nodeText_c", value = "")
    output$flowchart_c <- renderGrViz({
      dot_code <- generateDotCode_cochran()
      grViz(dot_code)
    })
  })
  
  output$downloadFlowchartPNG_c <- downloadHandler(
    filename = function() {
      paste("cochran_flowchart_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_cochran()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_png(tmp_file, file, width = 3000, height = 2000)
      unlink(tmp_file)
    }
  )
  
  output$downloadFlowchartSVG_c <- downloadHandler(
    filename = function() {
      paste("cochran_flowchart_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_cochran()
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(file)
    }
  )
  
  output$downloadFlowchartPDF_c <- downloadHandler(
    filename = function() {
      paste("cochran_flowchart_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      dot_code <- generateDotCode_cochran()
      tmp_file <- tempfile(fileext = ".svg")
      DiagrammeRsvg::export_svg(grViz(dot_code)) %>% writeLines(tmp_file)
      rsvg::rsvg_pdf(tmp_file, file)
      unlink(tmp_file)
    }
  )
  
  output$downloadCochranWord <- downloadHandler(
    filename = function() paste("Cochran_Sample_Size_Report_", Sys.Date(), ".docx", sep = ""),
    content = function(file) {
      res <- cochranSampleSize()
      p <- input$p
      z <- input$z
      e <- input$e_c
      N <- if (is.null(input$N_c) || is.na(input$N_c)) totalPopulation_c() else input$N_c
      alloc <- cochranAllocationData()
      
      doc <- read_docx() %>%
        body_add_par("CalcuStats", style = "heading 1") %>%
        body_add_par("Sample Size Calculation (Cochran's Method)", style = "heading 2") %>%
        body_add_par(paste("Estimated Proportion (p):", p), style = "Normal") %>%
        body_add_par(paste("Z-score (Z):", z), style = "Normal") %>%
        body_add_par(paste("Margin of Error (e):", e), style = "Normal") %>%
        body_add_par(paste("Non-response Rate:", input$non_response_c, "%"), style = "Normal")
      
      if (!is.null(N)) {
        doc <- doc %>% body_add_par(paste("Population Size (N):", N), style = "Normal")
      }
      
      doc <- doc %>%
        body_add_par(paste("Initial Sample Size (n₀):", ceiling(res$initial)), style = "Normal")
      
      if (res$finite_correction_applied) {
        doc <- doc %>%
          body_add_par(paste("Corrected Sample Size (n):", ceiling(res$corrected)), style = "Normal")
      }
      
      doc <- doc %>%
        body_add_par(paste("Adjusted Sample Size (accounting for non-response):", adjustedCochranSampleSize()), style = "Normal")
      
      if (!is.null(alloc)) {
        doc <- doc %>%
          body_add_par("Proportional Allocation Table:", style = "heading 2") %>%
          body_add_par("The proportional allocation formula is: n_i = (N_i / N) × n", style = "Normal")
        ft <- flextable(alloc) %>% autofit()
        doc <- doc %>% body_add_flextable(ft)
      }
      
      doc <- doc %>%
        body_add_par("Interpretation", style = "heading 2") %>%
        body_add_par(cochranInterpretationText(), style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  output$downloadCochranSteps <- downloadHandler(
    filename = function() {
      paste("Cochran_Calculation_Steps_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      res <- cochranSampleSize()
      p <- input$p
      z <- input$z
      e <- input$e_c
      N <- if (is.null(input$N_c) || is.na(input$N_c)) NULL else input$N_c
      alloc <- cochranAllocationData()
      
      steps <- c(
        "Cochran Sample Size Calculation Steps",
        "====================================",
        "",
        paste("Estimated Proportion (p):", p),
        paste("Z-score (Z):", z),
        paste("Margin of Error (e):", e),
        paste("Non-response Rate:", input$non_response_c, "%"),
        if (!is.null(N)) paste("Population Size (N):", N) else "Population Size: Infinite",
        "",
        "1. Basic Cochran Formula Calculation:",
        paste("Formula: n₀ = (Z² × p × (1-p)) / e²"),
        paste("Calculation: (", z, "² × ", p, " × ", (1-p), ") / ", e, "²"),
        paste("Numerator: ", z^2, " × ", p, " × ", (1-p), " = ", (z^2 * p * (1 - p))),
        paste("Denominator: ", e, "² = ", e^2),
        paste("Initial sample size (n₀): ", (z^2 * p * (1 - p)), " / ", e^2, " = ", res$initial)
      )
      
      if (res$finite_correction_applied) {
        steps <- c(steps,
                   "",
                   "2. Finite Population Correction:",
                   paste("Formula: n = n₀ / (1 + (n₀ - 1)/N)"),
                   paste("Calculation: ", res$initial, " / (1 + (", res$initial, " - 1)/", N, ")"),
                   paste("Denominator: 1 + (", res$initial, " - 1)/", N, " = ", (1 + (res$initial - 1)/N)),
                   paste("Corrected sample size: ", res$initial, " / ", (1 + (res$initial - 1)/N), " = ", res$corrected),
                   "",
                   paste("Final sample size before non-response adjustment: ", ceiling(res$corrected))
        )
      } else {
        steps <- c(steps,
                   "",
                   "2. No finite population correction applied (infinite population assumed)",
                   "",
                   paste("Final sample size before non-response adjustment: ", ceiling(res$initial))
        )
      }
      
      steps <- c(
        steps,
        "",
        "3. Non-response Adjustment:",
        paste("Non-response rate: ", input$non_response_c, "%"),
        paste("Adjusted sample size = ", ceiling(res$corrected), " / (1 - ", input$non_response_c/100, ")"),
        paste("Calculation: ", ceiling(res$corrected), " / ", (1 - input$non_response_c/100), " = ", ceiling(res$corrected) / (1 - input$non_response_c/100)),
        paste("Final adjusted sample size: ", adjustedCochranSampleSize())
      )
      
      if (!is.null(alloc)) {
        steps <- c(
          steps,
          "",
          "4. Proportional Allocation:",
          capture.output(print(alloc))
        )
      }
      
      interpretation <- cochranInterpretationText()
      
      steps <- c(steps, "", "Interpretation:", interpretation)
      
      writeLines(steps, file)
    }
  )
  
  # Power Analysis section
  powerResult <- reactiveVal()
  
  observeEvent(input$runPower, {
    result <- tryCatch({
      switch(input$testType,
             "Independent t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "two.sample"),
             "Paired t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "paired"),
             "One-sample t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "one.sample"),
             "One-Way ANOVA" = {
               # For One-Way ANOVA, k should be the number of groups (minimum 2)
               pwr.anova.test(f = input$effectSize, sig.level = input$alpha, power = input$power, k = 2)
             },
             "Two-Way ANOVA" = {
               # For Two-Way ANOVA, we need to calculate the degrees of freedom differently
               # This is a simplified approach - in practice, you'd need more parameters
               pwr.anova.test(f = input$effectSize, sig.level = input$alpha, power = input$power, k = 2)
             },
             "Proportion" = pwr.p.test(h = input$effectSize, sig.level = input$alpha, power = input$power),
             "Cororrelation" = pwr.r.test(r = input$effectSize, sig.level = input$alpha, power = input$power),
             "Chi-squared" = pwr.chisq.test(w = input$effectSize, sig.level = input$alpha, power = input$power, df = 1),
             "Simple Linear Regression" = pwr.f2.test(u = 1, f2 = input$effectSize, sig.level = input$alpha, power = input$power),
             "Multiple Linear Regression" = pwr.f2.test(u = input$predictors, f2 = input$effectSize, sig.level = input$alpha, power = input$power)
      )
    }, error = function(e) {
      return(paste("Error in power calculation:", e$message))
    })
    
    powerResult(result)
    output$powerResult <- renderPrint({ 
      if (is.character(result)) {
        cat(result)
      } else {
        result 
      }
    })
  })
  
  output$downloadPowerSteps <- downloadHandler(
    filename = function() paste("Power_Analysis_Steps_", Sys.Date(), ".txt", sep = ""),
    content = function(file) {
      req(powerResult())
      result <- powerResult()
      
      if (is.character(result)) {
        writeLines(c("Power Analysis Calculation Steps",
                     "===============================",
                     "",
                     result), file)
        return()
      }
      
      test_info <- switch(input$testType,
                          "Independent t-test" = "Independent samples t-test",
                          "Paired t-test" = "Paired samples t-test",
                          "One-sample t-test" = "One-sample t-test",
                          "One-Way ANOVA" = "One-Way ANOVA",
                          "Two-Way ANOVA" = "Two-Way ANOVA",
                          "Proportion" = "Test of proportion",
                          "Correlation" = "Correlation test",
                          "Chi-squared" = "Chi-squared test",
                          "Simple Linear Regression" = "Simple Linear Regression",
                          "Multiple Linear Regression" = "Multiple Linear Regression"
      )
      
      formula_text <- switch(input$testType,
                             "Independent t-test" = "pwr.t.test(d = effect_size, sig.level = alpha, power = power, type = 'two.sample')",
                             "Paired t-test" = "pwr.t.test(d = effect_size, sig.level = alpha, power = power, type = 'paired')",
                             "One-sample t-test" = "pwr.t.test(d = effect_size, sig.level = alpha, power = power, type = 'one.sample')",
                             "One-Way ANOVA" = "pwr.anova.test(f = effect_size, sig.level = alpha, power = power, k = number_of_groups)",
                             "Two-Way ANOVA" = "Note: Two-Way ANOVA requires more complex power calculation",
                             "Proportion" = "pwr.p.test(h = effect_size, sig.level = alpha, power = power)",
                             "Correlation" = "pwr.r.test(r = effect_size, sig.level = alpha, power = power)",
                             "Chi-squared" = "pwr.chisq.test(w = effect_size, sig.level = alpha, power = power, df = degrees_of_freedom)",
                             "Simple Linear Regression" = "pwr.f2.test(u = 1, f2 = effect_size, sig.level = alpha, power = power)",
                             "Multiple Linear Regression" = paste0("pwr.f2.test(u = ", input$predictors, ", f2 = effect_size, sig.level = alpha, power = power)")
      )
      
      steps <- c(
        "Power Analysis Calculation Steps",
        "===============================",
        "",
        paste("Test Type:", test_info),
        paste("Effect Size:", input$effectSize),
        paste("Significance Level (alpha):", input$alpha),
        paste("Desired Power:", input$power),
        if (input$testType == "Multiple Linear Regression") paste("Number of Predictors:", input$predictors) else "",
        "",
        "R Function Used:",
        formula_text,
        "",
        "Results:",
        capture.output(print(result)),
        "",
        "Interpretation:",
        paste("For a", test_info, "with an effect size of", input$effectSize, ","),
        paste("a significance level of", input$alpha, ","),
        paste("and desired power of", input$power, ","),
        if (input$testType %in% c("One-Way ANOVA", "Two-Way ANOVA")) {
          paste("the required sample size per group is approximately", ceiling(result$n), ".")
        } else if (input$testType == "Multiple Linear Regression") {
          paste("the required total sample size is", ceiling(result$v + input$predictors + 1), ".")
        } else {
          paste("the required sample size is", ceiling(result$n), ".")
        }
      )
      
      writeLines(steps, file)
    }
  )
  
  # Descriptive Statistics section
  descResult <- reactiveVal()
  
  observeEvent(input$runDesc, {
    req(input$dataInput)
    lines <- unlist(strsplit(input$dataInput, "[\n\r]+"))
    
    if (length(lines) < 2) {
      output$descResult <- renderPrint({ "Please include header (Variable name) and at least one row of data." })
      return()
    }
    
    header <- lines[1]
    data_lines <- lines[-1]
    
    suppressWarnings({
      nums <- as.numeric(data_lines)
      is_numeric <- !is.na(nums)
    })
    
    if (all(is_numeric)) {
      values <- nums[is_numeric]
      n_miss <- sum(is.na(nums))
      desc <- list(
        Mean = mean(values),
        SD = sd(values),
        Min = min(values),
        Max = max(values),
        Mode = as.numeric(names(sort(table(values), decreasing = TRUE)[1])),
        Median = median(values),
        Skewness = skewness(values),
        Kurtosis = kurtosis(values),
        Missing = n_miss,
        Shapiro_Wilk_p = shapiro.test(values)$p.value
      )
      desc$total <- sum(values)
      
      descResult(list(type = "numeric", header = header, results = desc))
      
      output$descResult <- renderPrint({
        cat(paste("Descriptive Statistics for:", header, "\n"))
        cat("(Numeric Data)\n\n")
        print(desc)
      })
    } else {
      categories <- trimws(data_lines)
      tbl <- table(factor(categories))
      cum_freq <- cumsum(tbl)
      percent <- prop.table(tbl) * 100
      n_miss <- sum(is.na(categories))
      
      df <- data.frame(
        Category = names(tbl),
        Frequency = as.vector(tbl),
        Percentage = round(as.vector(percent), 2),
        Cumulative_Frequency = as.vector(cum_freq)
      )
      
      df <- rbind(df, data.frame(
        Category = "Total", 
        Frequency = sum(tbl), 
        Percentage = 100, 
        Cumulative_Frequency = max(cum_freq))
      )
      
      descResult(list(type = "categorical", header = header, results = df, missing = n_miss))
      
      output$descResult <- renderPrint({
        cat(paste("Descriptive Statistics for:", header, "\n"))
        cat("(Categorical Data)\n\n")
        print(df)
        cat("\nMissing Values:", n_miss, "\n")
      })
    }
  })
  
  output$downloadDescSteps <- downloadHandler(
    filename = function() paste("Descriptive_Stats_Steps_", Sys.Date(), ".txt", sep = ""),
    content = function(file) {
      req(descResult())
      result <- descResult()
      
      if (result$type == "numeric") {
        steps <- c(
          "Descriptive Statistics Calculation Steps (Numeric Data)",
          "======================================================",
          "",
          paste("Variable Name:", result$header),
          "",
          "Calculated Statistics:",
          paste("Mean:", result$results$Mean),
          paste("Standard Deviation:", result$results$SD),
          paste("Minimum:", result$results$Min),
          paste("Maximum:", result$results$Max),
          paste("Median:", result$results$Median),
          paste("Mode:", result$results$Mode),
          paste("Skewness:", result$results$Skewness),
          paste("Kurtosis:", result$results$Kurtosis),
          paste("Total:", result$results$total),
          paste("Missing Values:", result$results$Missing),
          paste("Shapiro-Wilk p-value:", result$results$Shapiro_Wilk_p),
          "",
          "Interpretation:",
          ifelse(result$results$Shapiro_Wilk_p > 0.05, 
                 "The data appears to be normally distributed", 
                 "The data does not appear to be normally distributed"),
          paste("(Shapiro-Wilk test p-value:", result$results$Shapiro_Wilk_p, ")")
        )
      } else {
        steps <- c(
          "Descriptive Statistics Calculation Steps (Categorical Data)",
          "==========================================================",
          "",
          paste("Variable Name:", result$header),
          "",
          "Frequency Table:",
          capture.output(print(result$results)),
          "",
          paste("Missing Values:", result$missing)
        )
      }
      
      writeLines(steps, file)
    }
  )
}

# Add JavaScript for localStorage functionality
jscode <- "
// Function to save values to localStorage
shinyjs.saveValues = function(params) {
  localStorage.setItem(params.section, params.values);
}

// Function to clear values from localStorage
shinyjs.clearValues = function(params) {
  localStorage.removeItem(params);
}

// Function to initialize values from localStorage
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
