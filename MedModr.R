library(shiny)
library(lavaan)
library(semPlot)
library(ggplot2)
library(readr)
library(readxl)
library(haven)
library(DT)
library(colourpicker)
library(officer)
library(magick)
library(flextable)
library(shinythemes)
library(shinyjs)
library(shinydashboard)

app_name <- "MedModr"
app_version <- "3.0.0"
release_date <- "October 2025"

ui <- fluidPage(
  useShinyjs(),
  title = "MedModr | Mediation and Moderation Analyses Software",
  titleWidth = 350,
  
  theme = shinytheme("cosmo"),
  tags$head(
    tags$style(HTML("
  @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&family=Poppins:wght@400;500;600;700&display=swap');
  
  :root {
    --primary-blue: #3498db;
    --primary-purple: #9b59b6;
    --primary-green: #2ecc71;
    --dark-blue: #2980b9;
    --dark-purple: #8e44ad;
    --dark-green: #27ae60;
    --dark-gray: #2c3e50;
    --medium-gray: #34495e;
    --light-gray: #ecf0f1;
    --background-gray: #f8f9fa;
    --border-gray: #e0e0e0;
    --text-dark: #2c3e50;
    --text-medium: #7f8c8d;
    --text-light: #95a5a6;
    --success-light: #d5f4e6;
    --info-light: #e8f4fc;
    --warning-light: #fff3cd;
  }
  
  body {
    font-family: 'Roboto', sans-serif;
    background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    background-attachment: fixed;
    color: var(--text-dark);
    line-height: 1.6;
  }
  
  .app-title {
    font-family: 'Poppins', sans-serif;
    font-weight: 700;
    font-size: 2.5rem;
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    text-align: center;
    margin: 1rem 0 2rem 0;
    padding: 1rem;
    text-shadow: 0 2px 4px rgba(0,0,0,0.1);
  }
  
  .app-subtitle {
    font-family: 'Poppins', sans-serif;
    font-size: 0.9rem;
    color: var(--text-medium);
    text-align: center;
    margin-top: -1.5rem;
    margin-bottom: 2rem;
    font-weight: 400;
  }
  
  .interface-choice-card {
    background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 15px;
    padding: 30px;
    margin: 20px;
    text-align: center;
    cursor: pointer;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border: 2px solid transparent;
    box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
    height: 200px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
  }
  
  .interface-choice-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 35px rgba(0, 0, 0, 0.15);
    border-color: var(--primary-blue);
  }
  
  .interface-choice-card.selected {
    border-color: var(--primary-blue);
    background: linear-gradient(135deg, #e8f4fc 0%, #f8f9fa 100%);
  }
  
  .choice-icon {
    font-size: 3rem;
    margin-bottom: 15px;
    color: var(--primary-blue);
  }
  
  .choice-title {
    font-size: 1.3rem;
    font-weight: 600;
    margin-bottom: 10px;
    color: var(--dark-gray);
  }
  
  .choice-description {
    font-size: 0.9rem;
    color: var(--text-medium);
    line-height: 1.4;
  }
  
  .variable-selection-box {
    background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 10px;
    padding: 20px;
    margin: 15px 0;
    border: 2px dashed var(--border-gray);
    transition: all 0.3s ease;
    min-height: 80px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    cursor: pointer;
  }
  
  .variable-selection-box:hover {
    border-color: var(--primary-blue);
    background: linear-gradient(135deg, #e8f4fc 0%, #f8f9fa 100%);
  }
  
  .variable-selection-box.has-variable {
    border: 2px solid var(--primary-green);
    background: linear-gradient(135deg, var(--success-light) 0%, #f8f9fa 100%);
  }
  
  .variable-label {
    font-weight: 600;
    color: var(--dark-gray);
    margin-bottom: 5px;
  }
  
  .variable-name {
    font-weight: 500;
    color: var(--primary-blue);
    font-size: 1.1rem;
  }
  
  .click-instruction {
    font-size: 0.8rem;
    color: var(--text-light);
    font-style: italic;
  }
  
  .navbar {
    background: linear-gradient(135deg, var(--dark-gray) 0%, var(--medium-gray) 100%) !important;
    border: none !important;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
  }
  
  .sidebar {
    background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 15px;
    box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
    padding: 25px;
    margin-right: 15px;
    height: 90vh;
    overflow-y: auto;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border: 1px solid rgba(255,255,255,0.2);
  }
  
  .sidebar:hover {
    box-shadow: 0 12px 35px rgba(0, 0, 0, 0.15);
    transform: translateY(-2px);
  }
  
  .main-panel {
    background: linear-gradient(180deg, #ffffff 0%, #fafbfc 100%);
    border-radius: 15px;
    box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
    padding: 25px;
    height: 90vh;
    overflow-y: auto;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border: 1px solid rgba(255,255,255,0.2);
  }
  
  .well {
    background: linear-gradient(135deg, #ffffff 0%, var(--light-gray) 100%);
    border-radius: 12px;
    border: 1px solid var(--border-gray);
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    transition: all 0.3s ease;
  }
  
  .well:hover {
    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.12);
  }
  
  .btn-primary {
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--dark-blue) 100%);
    border: none;
    color: white;
    font-weight: 500;
    border-radius: 8px;
    padding: 10px 20px;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
  }
  
  .btn-primary:hover {
    background: linear-gradient(135deg, var(--dark-blue) 0%, #2472a4 100%);
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(52, 152, 219, 0.4);
    border: none;
  }
  
  .btn-success {
    background: linear-gradient(135deg, var(--primary-green) 0%, var(--dark-green) 100%);
    border: none;
    border-radius: 8px;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    box-shadow: 0 4px 12px rgba(46, 204, 113, 0.3);
  }
  
  .btn-success:hover {
    background: linear-gradient(135deg, var(--dark-green) 0%, #219653 100%);
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(46, 204, 113, 0.4);
    border: none;
  }
  
  .btn-info {
    background: linear-gradient(135deg, var(--primary-purple) 0%, var(--dark-purple) 100%);
    border: none;
    border-radius: 8px;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    box-shadow: 0 4px 12px rgba(155, 89, 182, 0.3);
  }
  
  .btn-info:hover {
    background: linear-gradient(135deg, var(--dark-purple) 0%, #7d3c98 100%);
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(155, 89, 182, 0.4);
    border: none;
  }
  
  .tab-content {
    background: linear-gradient(180deg, #ffffff 0%, #fafbfc 100%);
    border-radius: 0 0 12px 12px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    padding: 25px;
    border: 1px solid var(--border-gray);
    border-top: none;
  }
  
  .nav-tabs > li > a {
    color: var(--text-medium);
    font-weight: 500;
    border-radius: 8px 8px 0 0;
    transition: all 0.3s ease;
    margin-right: 5px;
  }
  
  .nav-tabs > li.active > a,
  .nav-tabs > li.active > a:hover,
  .nav-tabs > li.active > a:focus {
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
    color: white;
    border: none;
    box-shadow: 0 -4px 12px rgba(52, 152, 219, 0.3);
  }
  
  .nav-tabs > li > a:hover {
    background: linear-gradient(135deg, var(--light-gray) 0%, #e9ecef 100%);
    color: var(--primary-blue);
    border: none;
  }
  
  .instruction-box {
    background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%);
    border-left: 4px solid var(--primary-blue);
    padding: 20px;
    margin-bottom: 25px;
    border-radius: 10px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  }
  
  .interpretation-box {
    background: linear-gradient(135deg, var(--success-light) 0%, #e8f5e8 100%);
    border-left: 4px solid var(--primary-green);
    padding: 20px;
    margin: 25px 0;
    border-radius: 10px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  }
  
  .analysis-step {
    background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 10px;
    padding: 20px;
    margin-bottom: 20px;
    border-left: 4px solid var(--primary-purple);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
    transition: all 0.3s ease;
  }
  
  .analysis-step:hover {
    transform: translateX(5px);
    box-shadow: 0 6px 18px rgba(0, 0, 0, 0.08);
  }
  
  h3, h4 {
    font-weight: 600;
    color: var(--text-dark);
  }
  
  h3 {
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    padding-bottom: 10px;
    margin-bottom: 20px;
    border-bottom: 2px solid var(--border-gray);
  }
  
  h4 {
    color: var(--medium-gray);
    font-weight: 500;
  }
  
  select, input, textarea {
    border-radius: 8px !important;
    border: 1px solid var(--border-gray) !important;
    transition: all 0.3s ease;
    padding: 10px 12px;
  }
  
  select:focus, input:focus, textarea:focus {
    border-color: var(--primary-blue) !important;
    box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1) !important;
    outline: none;
  }
  
  .file-input {
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    border-radius: 8px;
    border: 2px dashed #bdc3c7;
    transition: all 0.3s ease;
    padding: 15px;
  }
  
  .file-input:hover {
    border-color: var(--primary-blue);
    background: linear-gradient(135deg, #e3f2fd 0%, #f8f9fa 100%);
  }
  
  .data-summary {
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    padding: 20px;
    border-radius: 10px;
    border: 1px solid var(--border-gray);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  }
  
  .scrollable-section {
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    border-radius: 10px;
    border: 1px solid var(--border-gray);
    box-shadow: inset 0 2px 4px rgba(0,0,0,0.05);
  }
  
  .diagram-controls-left {
    background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    border: 1px solid var(--border-gray);
  }
  
  .bootstrap-info {
    background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%);
    border-left: 4px solid var(--primary-blue);
    padding: 15px;
    margin: 15px 0;
    border-radius: 8px;
  }
  
  .diagram-title-container {
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
    border-radius: 10px;
    padding: 15px;
    margin-bottom: 20px;
    box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
  }
  
  .diagram-title {
    font-size: 1.5rem;
    font-weight: 600;
    color: white;
    margin: 0;
    text-shadow: 0 2px 4px rgba(0,0,0,0.1);
  }
  
  .resizable-text {
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    border: 1px solid var(--border-gray);
    border-radius: 8px;
    box-shadow: inset 0 2px 4px rgba(0,0,0,0.05);
  }
  
  /* Custom scrollbar */
  ::-webkit-scrollbar {
    width: 8px;
  }
  
  ::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 4px;
  }
  
  ::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
    border-radius: 4px;
  }
  
  ::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(135deg, var(--dark-blue) 0%, var(--dark-purple) 100%);
  }
  
  /* Loading animation */
  .shiny-notification {
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
    color: white;
    border-radius: 10px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    border: none;
  }
  
  /* Table styling */
  .dataTables_wrapper {
    background: white;
    border-radius: 10px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
  }
  
  /* Card-like sections */
  .results-content > div {
    background: white;
    border-radius: 10px;
    padding: 20px;
    margin-bottom: 20px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    border: 1px solid var(--border-gray);
  }
  
  .sidebar-collapsed {
    display: none !important;
  }
  
  .main-panel-expanded {
    width: 100% !important;
  }
  
  /* Ensure diagram controls remain visible and properly positioned */
  .diagram-controls-left {
    background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    border: 1px solid var(--border-gray);
    padding: 15px;
    height: fit-content;
    position: sticky;
    top: 20px;
  }
  
  .diagram-main-area {
    height: 100%;
    display: flex;
    flex-direction: column;
  }
  
  /* Smooth transitions for sidebar */
  #sidebar {
    transition: all 0.3s ease-in-out;
  }
  
  #main_panel {
    transition: all 0.3s ease-in-out;
  }
  
  /* Diagram Controls - Scrollable */
  .diagram-controls-left {
    background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    border: 1px solid var(--border-gray);
    padding: 20px;
    height: 80vh;
    overflow-y: auto;
    position: sticky;
    top: 20px;
    margin-bottom: 20px;
  }
  
  /* Diagram Main Area - Static */
  .diagram-main-area {
    background: white;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    height: 85vh;
    overflow: hidden;
    position: relative;
  }
  
  /* Ensure the diagram itself doesn't create scroll */
  #semPlot {
    width: 100% !important;
    height: 100% !important;
    max-height: 100% !important;
  }
  
  /* Custom scrollbar for diagram controls */
  .diagram-controls-left::-webkit-scrollbar {
    width: 6px;
  }
  
  .diagram-controls-left::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 3px;
  }
  
  .diagram-controls-left::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
    border-radius: 3px;
  }
  
  .diagram-controls-left::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(135deg, var(--dark-blue) 0%, var(--dark-purple) 100%);
  }
  
  /* Responsive adjustments */
  @media (max-height: 800px) {
    .diagram-controls-left {
      height: 75vh;
    }
    
    .diagram-main-area {
      height: 80vh;
    }
  }
  
  @media (max-height: 600px) {
    .diagram-controls-left {
      height: 70vh;
    }
    
    .diagram-main-area {
      height: 75vh;
    }
  }
  
  /* Analysis options styling */
  .analysis-options {
    background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 10px;
    padding: 20px;
    margin: 15px 0;
    border: 1px solid var(--border-gray);
  }
  
  .option-group {
    margin: 15px 0;
  }
  
  .option-label {
    font-weight: 600;
    color: var(--dark-gray);
    margin-bottom: 8px;
  }
  
"))
  ),
  
  titlePanel(
    div(
      class = "app-title", 
      HTML(paste0(app_name, " <small style='font-size: 14px; color: #7f8c8d;'>v", app_version, " (", release_date, ")</small>"))
    )
  ),
  
  # Interface Selection Screen
  conditionalPanel(
    condition = "output.interface_selected == false",
    fluidRow(
      column(12, align = "center",
             h3("Choose Your Analysis Interface", style = "margin-bottom: 30px;"),
             p("Select how you want to interact with the app:", style = "margin-bottom: 40px; font-size: 1.1rem; color: var(--text-medium);")
      )
    ),
    fluidRow(
      column(6,
             div(class = "interface-choice-card", id = "interactive_choice",
                 div(class = "choice-icon", icon("mouse-pointer")),
                 div(class = "choice-title", "Interactive Interface"),
                 div(class = "choice-description", 
                     "Point-and-click interface for easy variable selection. Perfect for beginners and quick analyses.")
             )
      ),
      column(6,
             div(class = "interface-choice-card", id = "syntax_choice",
                 div(class = "choice-icon", icon("code")),
                 div(class = "choice-title", "Syntax Interface"),
                 div(class = "choice-description", 
                     "Write lavaan syntax directly for maximum flexibility and control. For R users.")
             )
      )
    ),
    fluidRow(
      column(12, align = "center", style = "margin-top: 40px;",
             actionButton("confirm_interface", "Continue with Selected Interface", 
                          class = "btn-primary btn-lg", icon = icon("arrow-right"))
      )
    )
  ),
  
  # Main App Interface (hidden until interface is selected)
  conditionalPanel(
    condition = "output.interface_selected == true",
    sidebarLayout(
      sidebarPanel(
        width = 4,
        id = "sidebar",
        class = "sidebar",
        
        # Interface switcher
        div(class = "instruction-box",
            fluidRow(
              column(8,
                     h4("Current Interface:", style = "margin: 0;"),
                     textOutput("current_interface", inline = TRUE)
              ),
              column(4, align = "right",
                     actionButton("switch_interface", "Switch Interface", 
                                  class = "btn-info btn-sm", icon = icon("exchange"))
              )
            )
        ),
        
        # Data Input Section
        h4("Data Input", icon("database")),
        fileInput("datafile", "Upload your data (CSV, XLSX, Stata, or SPSS):", 
                  accept = c(".csv", ".xlsx", ".xls", ".dta", ".sav", ".zsav", ".por")),
        
        uiOutput("var_preview"),
        
        # Interactive Interface UI
        conditionalPanel(
          condition = "output.current_interface == 'interactive'",
          h4("Model Specification", icon("sliders")),
          selectInput("analysis_type_interactive", "Select Analysis Type:",
                      choices = c("Simple Mediation", "Serial Mediation", "Moderation")),
          
          # Variable Selection Boxes
          uiOutput("variable_selection_ui"),
          
          # Analysis Options
          h4("Analysis Options", icon("cogs")),
          # Removed Use Constructs checkbox from interactive interface
          
          div(class = "analysis-options",
              h5("Output Options:"),
              conditionalPanel(
                condition = "input.analysis_type_interactive == 'Moderation'",
                checkboxGroupInput("output_options_moderation", "Select output to include:",
                                   choices = c("Interaction Effects" = "interaction",
                                               "Simple Slopes" = "slopes",
                                               "Model Fit Indices" = "fit"),
                                   selected = c("interaction", "slopes", "fit"))
              ),
              conditionalPanel(
                condition = "input.analysis_type_interactive != 'Moderation'",
                checkboxGroupInput("output_options", "Select output to include:",
                                   choices = c("Total Effects" = "total",
                                               "Direct Effects" = "direct",
                                               "Indirect Effects" = "indirect",
                                               "Path Coefficients" = "paths",
                                               "Model Fit Indices" = "fit"),
                                   selected = c("total", "direct", "indirect", "paths", "fit"))
              )
          )
        ),
        
        # Syntax Interface UI
        conditionalPanel(
          condition = "output.current_interface == 'syntax'",
          h4("Model Specification", icon("code")),
          div(class = "instruction-box",
              p("Tip: Check the 'Variables in Data' tab to see your variable names before writing your model."),
              p("This app uses the", tags$strong("lavaan"), "package for structural equation modeling analysis.")
          ),
          textAreaInput("model_syntax", "Specify lavaan Model Syntax:", 
                        placeholder = "Enter your model syntax here...", height = "200px"),
          
          h4("Analysis Settings", icon("cogs")),
          selectInput("analysis_type", "Select Analysis Type:",
                      choices = c("Simple Mediation", "Serial Mediation"))
          # Removed Use Constructs checkbox from syntax interface as well
        ),
        
        # Common Analysis Settings
        div(style = "display: none;",
            checkboxInput("use_bootstrap", "Use Bootstrap Confidence Intervals", value = FALSE),
            conditionalPanel(
              condition = "input.use_bootstrap == true",
              numericInput("bootstrap_samples", "Number of Bootstrap Samples:", 
                           value = 5000, min = 1000, max = 10000, step = 1000),
              div(class = "bootstrap-info",
                  p("Note: By default, the app shows results without bootstrap. If you want bootstrap results, check this option and specify the number of samples.")
              )
            )
        ),
        
        uiOutput("covariate_ui"),
        
        actionButton("run", "Run Analysis", icon = icon("play"), 
                     class = "btn-primary btn-lg"),
        br(), br(),
        downloadButton("download_report", "Download Word Report", class = "btn-success"),
        
        
        # Interactive Interface Guide
        conditionalPanel(
          condition = "output.current_interface == 'interactive'",
          div(class = "scrollable-section",
              h4("Interactive Guide:", icon("info-circle")),
              uiOutput("interactive_guide")
          )
        ),
        
        # Syntax Interface Guide
        conditionalPanel(
          condition = "output.current_interface == 'syntax'",
          div(class = "scrollable-section",
              conditionalPanel(
                condition = "input.analysis_type == 'Simple Mediation'",
                h4("Simple Mediation Guide:", icon("info-circle")),
                tags$ul(
                  tags$li("X = Independent Variable"),
                  tags$li("M = Mediator Variable"),
                  tags$li("Y = Dependent Variable"),
                  tags$li("a = X → M path"),
                  tags$li("b = M → Y path (controlling for X)"),
                  tags$li("c' = Direct effect of X → Y (controlling for M)")
                ),
                h4("Example Syntax:"),
                verbatimTextOutput("simple_mediation_syntax")
              ),
              
              conditionalPanel(
                condition = "input.analysis_type == 'Serial Mediation'",
                h4("Serial Mediation Guide:", icon("info-circle")),
                tags$ul(
                  tags$li("X = Independent Variable"),
                  tags$li("M1 = Mediator 1"),
                  tags$li("M2 = Mediator 2"),
                  tags$li("Y = Dependent Variable"),
                  tags$li("a1 = X → M1 path"),
                  tags$li("a2 = M1 → M2 path"),
                  tags$li("b1 = M1 → Y path"),
                  tags$li("b2 = M2 → Y path"),
                  tags$li("d = X → M2 path (direct)"),
                  tags$li("c = Direct effect of X → Y")
                ),
                h4("Example Syntax:"),
                verbatimTextOutput("serial_mediation_syntax")
              )
          )
        )
      ),
      
      mainPanel(
        width = 8,
        id = "main_panel",
        class = "main-panel",
        tabsetPanel(
          id = "main_tabs",
          tabPanel("Results", 
                   div(class = "results-content",
                       uiOutput("fitUI"),
                       uiOutput("mediationTableUI"),
                       uiOutput("regressionTableUI"),
                       uiOutput("effectsUI"),
                       h3("Results Interpretation", icon("comment")),
                       uiOutput("mediationInterpretation")
                   )),
          tabPanel("Diagram", 
                   div(class = "diagram-full-view",
                       fluidRow(
                         column(
                           width = 3,
                           class = "diagram-controls-left",
                           id = "diagram_controls",
                           h4("Visualization Controls", icon("chart-line")),
                           actionButton("toggle_sidebar", "Show/Hide Data Panel", 
                                        icon = icon("bars"), 
                                        class = "btn-info btn-sm"),
                           br(), br(),
                           
                           # Tab for switching between path diagram and slope plot
                           tabsetPanel(
                             id = "viz_tabs",
                             type = "pills",
                             
                             # Path Diagram Tab
                             tabPanel(
                               "Path Diagram",
                               icon = icon("project-diagram"),
                               
                               h5("Layout Options"),
                               textInput("diagram_title", "Diagram Title:", 
                                         placeholder = "Enter diagram title"),
                               selectInput("diagram_layout", "Layout:",
                                           choices = c("tree", "circle", "spring", "tree2", "tree3"),
                                           selected = "tree"),
                               
                               h5("Sizing Options"),
                               sliderInput("diagram_width", "Download Width:", 
                                           min = 800, max = 3000, value = 1600, step = 100),
                               sliderInput("diagram_height", "Download Height:", 
                                           min = 600, max = 2000, value = 1200, step = 100),
                               
                               h5("Node Options"),
                               sliderInput("node_size", "Node Size:", 
                                           min = 5, max = 20, value = 10, step = 1),
                               sliderInput("text_size", "Node Text Size:", 
                                           min = 0.5, max = 3, value = 1, step = 0.1),
                               
                               h5("Edge Options"),
                               sliderInput("edge_label_size", "Edge Label Size:", 
                                           min = 0.5, max = 2, value = 1.2, step = 0.1),
                               sliderInput("arrow_size", "Arrow Size:", 
                                           min = 0.5, max = 2, value = 1, step = 0.1),
                               sliderInput("edge_width", "Path Line Width:", 
                                           min = 0.5, max = 5, value = 1.5, step = 0.5),
                               
                               h5("Color Options"),
                               colourpicker::colourInput("man_color", "Observed Variable Color:", value = "lightyellow"),
                               colourpicker::colourInput("lat_color", "Latent Variable Color:", value = "skyblue"),
                               colourpicker::colourInput("edge_color", "Path Line Color:", value = "black"),
                               
                               h5("Node Labels"),
                               textInput("node_x_label", "X Variable Label:", 
                                         placeholder = "Label for X variable"),
                               textInput("node_y_label", "Y Variable Label:", 
                                         placeholder = "Label for Y variable"),
                               textInput("node_m_label", "M Variable Label:", 
                                         placeholder = "Label for M variable"),
                               textInput("node_m1_label", "M1 Variable Label:", 
                                         placeholder = "Label for M1 variable"),
                               
                               downloadButton("download_diagram", "Download Diagram (PNG)", 
                                              class = "btn-success btn-block")
                             ),
                             
                             # Slope Plot Tab (only visible for moderation analysis)
                             tabPanel(
                               "Slope Plot",
                               icon = icon("line-chart"),
                               conditionalPanel(
                                 condition = "output.is_moderation_analysis == true",
                                 
                                 div(class = "instruction-box",
                                     h5(icon("info-circle"), "Simple Slopes Plot"),
                                     p("This plot shows the relationship between the independent variable and dependent variable at different levels of the moderator."),
                                     p("The slopes represent the effect of X on Y at low (-1 SD), average, and high (+1 SD) levels of the moderator.")
                                 ),
                                 
                                 h5("Plot Options"),
                                 
                                 # Moderator level labels
                                 textInput("mod_low_label", "Low Moderator Label:", 
                                           value = "Low Moderator (-1 SD)",
                                           placeholder = "Label for low moderator level"),
                                 textInput("mod_avg_label", "Average Moderator Label:", 
                                           value = "Average Moderator (Mean)",
                                           placeholder = "Label for average moderator level"),
                                 textInput("mod_high_label", "High Moderator Label:", 
                                           value = "High Moderator (+1 SD)",
                                           placeholder = "Label for high moderator level"),
                                 
                                 hr(),
                                 
                                 h5("Aesthetic Options"),
                                 colourpicker::colourInput("slope_low_color", "Low Moderator Line Color:", value = "#E74C3C"),
                                 colourpicker::colourInput("slope_avg_color", "Average Moderator Line Color:", value = "#2ECC71"),
                                 colourpicker::colourInput("slope_high_color", "High Moderator Line Color:", value = "#3498DB"),
                                 
                                 sliderInput("slope_line_width", "Line Width:", 
                                             min = 0.5, max = 3, value = 1.5, step = 0.1),
                                 sliderInput("slope_point_size", "Data Point Size:", 
                                             min = 0.5, max = 5, value = 2, step = 0.1),
                                 sliderInput("slope_alpha", "Transparency:", 
                                             min = 0.2, max = 1, value = 0.6, step = 0.05),
                                 
                                 hr(),
                                 
                                 h5("Axis Labels"),
                                 textInput("x_axis_label", "X-Axis Label:", 
                                           value = "Independent Variable (X)",
                                           placeholder = "Label for X axis"),
                                 textInput("y_axis_label", "Y-Axis Label:", 
                                           value = "Dependent Variable (Y)",
                                           placeholder = "Label for Y axis"),
                                 
                                 hr(),
                                 
                                 h5("Download Options"),
                                 sliderInput("slope_plot_width", "Plot Width (pixels):", 
                                             min = 400, max = 1200, value = 800, step = 50),
                                 sliderInput("slope_plot_height", "Plot Height (pixels):", 
                                             min = 300, max = 900, value = 600, step = 50),
                                 
                                 downloadButton("download_slope_plot", "Download Slope Plot (PNG)", 
                                                class = "btn-info btn-block"),
                                 
                                 br(),
                                 
                                 # Simple slopes interpretation
                                 div(class = "bootstrap-info",
                                     h5(icon("calculator"), "Simple Slopes Values"),
                                     verbatimTextOutput("simple_slopes_values")
                                 )
                               ),
                               
                               # Message when moderation analysis not selected
                               conditionalPanel(
                                 condition = "output.is_moderation_analysis == false",
                                 div(class = "well", align = "center", style = "padding: 40px;",
                                     icon("info-circle", class = "fa-3x", style = "color: #3498db; margin-bottom: 15px;"),
                                     h4("Slope Plot Available for Moderation Analysis Only"),
                                     p("Please select 'Moderation' as your analysis type to view and download the simple slopes plot."),
                                     p("The slope plot visualizes the effect of the independent variable on the dependent variable at different levels of the moderator.")
                                 )
                               )
                             )
                           )
                         ),
                         
                         column(
                           width = 9,
                           div(class = "diagram-main-area",
                               uiOutput("diagram_title_ui"),
                               conditionalPanel(
                                 condition = "input.viz_tabs == 'Path Diagram'",
                                 plotOutput("semPlot", height = "100%")
                               ),
                               conditionalPanel(
                                 condition = "input.viz_tabs == 'Slope Plot'",
                                 plotOutput("slopePlot", height = "100%")
                               )
                           )
                         )
                         
                       )
                   )),
          
          tabPanel("Variables in Data", 
                   div(class = "instruction-box",
                       h4("Important:"),
                       p("Use the variable names shown below when writing your model syntax.")
                   ),
                   DTOutput("var_table"),
                   br(),
                   h4("Data Summary"),
                   verbatimTextOutput("data_summary")),
          
          tabPanel("How to Run Analysis",
                   div(class = "well",
                       h3("Step-by-Step Analysis Guide")),
                   
                   div(class = "instruction-box",
                       h4("How to perform simple mediation:"),
                       div(class = "analysis-step",
                           h5("Step 1: Use this syntax:"),
                           verbatimTextOutput("simple_mediation_howto1")),
                       div(class = "analysis-step",
                           h5("Step 2: Replace variables in the syntax"),
                           p("Copy this syntax into the Model Specification Box, replace only X, Y, and M with your variables"),
                           tags$ul(
                             tags$li("X = Independent Variable"),
                             tags$li("M = Mediator Variable"),
                             tags$li("Y = Dependent Variable")
                           ),
                           h5("Example:"),
                           verbatimTextOutput("simple_mediation_example")),
                       div(class = "analysis-step",
                           h5("Step 3: Add indirect effect syntax"),
                           p("Copy and add this syntax to your model:"),
                           verbatimTextOutput("simple_mediation_indirect"))
                   ),
                   
                   div(class = "instruction-box",
                       h4("How to perform serial mediation:"),
                       div(class = "analysis-step",
                           h5("Step 1: Use this syntax:"),
                           verbatimTextOutput("serial_mediation_howto1")),
                       div(class = "analysis-step",
                           h5("Step 2: Replace variables in the syntax"),
                           p("Copy this syntax into the Model Specification Box, replace X, Y, M1, and M2 with your variables"),
                           tags$ul(
                             tags$li("X = Independent Variable"),
                             tags$li("M1 = Mediator 1"),
                             tags$li("M2 = Mediator 2"),
                             tags$li("Y = Dependent Variable")
                           ),
                           h5("Example:"),
                           verbatimTextOutput("serial_mediation_example")),
                       div(class = "analysis-step",
                           h5("Step 3: Add indirect effects syntax"),
                           p("Copy and add this syntax to your model:"),
                           verbatimTextOutput("serial_mediation_indirect"))
                   ),
                   
                   div(class = "instruction-box",
                       h4("How to perform moderation:"),
                       div(class = "analysis-step",
                           h5("Step 1: Use this syntax:"),
                           verbatimTextOutput("moderation_howto1")),
                       div(class = "analysis-step",
                           h5("Step 2: Replace variables in the syntax"),
                           p("Copy this syntax into the Model Specification Box, replace X, Y, W, and XW with your variables"),
                           tags$ul(
                             tags$li("X = Independent Variable"),
                             tags$li("W = Moderator Variable"),
                             tags$li("XW = Interaction Term (X × W)"),
                             tags$li("Y = Dependent Variable")
                           ),
                           h5("Example:"),
                           verbatimTextOutput("moderation_example")),
                       div(class = "analysis-step",
                           h5("Step 3: Add simple slopes syntax"),
                           p("Copy and add this syntax to your model:"),
                           verbatimTextOutput("moderation_slopes"))
                   )
          ),
          tabPanel("About",
                   div(class = "well",
                       h4(app_name, "v", app_version),
                       p("A user-friendly interface for mediation, serial mediation, and moderation analysis using the lavaan package for structural equation modeling."),
                       h5("Key Features:"),
                       tags$ul(
                         tags$li("Dual interface: Interactive point-and-click AND syntax-based"),
                         tags$li("Uses lavaan package for robust statistical analysis"),
                         tags$li("Supports simple and serial mediation models"),
                         tags$li("Moderation analysis with interaction terms"),
                         tags$li("Interactive and customizable path diagrams"),
                         tags$li("Professional report generation"),
                         tags$li("Adjust for confounding variables")
                       ),
                       h5("Statistical Engine:"),
                       p("This app uses the", tags$strong("lavaan"), "package (Latent Variable Analysis) for all structural equation modeling analyses."),
                       h5("Developed by:", style = "margin-bottom: 5px;"),
                       div(style = "display: flex; align-items: center; gap: 15px; margin-bottom: 15px;",
                           p("Mudasir Mohammed Ibrahim", style = "margin: 0; font-weight: 500;"),
                           a(href = "https://scholar.google.com/citations?user=xEFzAvgAAAAJ&hl=en", 
                             target = "_blank",
                             icon("graduation-cap", class = "fa-2x", style = "color: #4285f4; transition: all 0.3s ease;"),
                             style = "text-decoration: none;"),
                           a(href = "https://orcid.org/0000-0002-9049-8222", 
                             target = "_blank",
                             icon("orcid", class = "fa-2x", style = "color: #A6CE39; transition: all 0.3s ease;"),
                             style = "text-decoration: none;"),
                           a(href = "https://www.researchgate.net/profile/Mudasir-Ibrahim", 
                             target = "_blank",
                             icon("researchgate", class = "fa-2x", style = "color: #00CCBB; transition: all 0.3s ease;"),
                             style = "text-decoration: none;")
                       ),
                       h5("Contact:"),
                       p("mudassiribrahim30@gmail.com"),
                       h5("License:"),
                       p("MIT License - Free for academic and research use")
                   ))
        )
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Reactive values for interface management
  interface_selected <- reactiveVal(FALSE)
  current_interface <- reactiveVal("") # "interactive" or "syntax"
  selected_variables <- reactiveValues(
    X = NULL, Y = NULL, M = NULL, M1 = NULL, M2 = NULL, W = NULL
  )
  
  # Output for interface selection condition
  output$interface_selected <- reactive({
    interface_selected()
  })
  outputOptions(output, "interface_selected", suspendWhenHidden = FALSE)
  
  # Output for current interface display
  output$current_interface <- renderText({
    current_interface()
  })
  
  # Create a reactive value to track moderation analysis
  is_moderation <- reactiveVal(FALSE)
  
  # Observe and update when analysis type changes
  observe({
    if (interface_selected()) {
      if (current_interface() == "interactive") {
        if (!is.null(input$analysis_type_interactive)) {
          is_moderation(input$analysis_type_interactive == "Moderation")
        } else {
          is_moderation(FALSE)
        }
      } else if (current_interface() == "syntax") {
        syntax <- input$model_syntax
        if (!is.null(syntax) && nchar(syntax) > 0) {
          has_interaction <- grepl("\\*", syntax) && grepl("~", syntax)
          has_slopes <- grepl("Simple_Slope", syntax)
          is_moderation(has_interaction || has_slopes)
        } else {
          is_moderation(FALSE)
        }
      }
    } else {
      is_moderation(FALSE)
    }
  })
  
  # Output for UI condition
  output$is_moderation_analysis <- reactive({
    is_moderation()
  })
  outputOptions(output, "is_moderation_analysis", suspendWhenHidden = FALSE)
  
  
  # Initialize interface selection
  observe({
    runjs("
      $('#interactive_choice').click(function() {
        Shiny.setInputValue('interactive_choice', Math.random());
      });
      $('#syntax_choice').click(function() {
        Shiny.setInputValue('syntax_choice', Math.random());
      });
    ")
  })
  
  # Handle interface choice clicks
  observeEvent(input$interactive_choice, {
    removeClass(selector = "#syntax_choice", class = "selected")
    addClass(id = "interactive_choice", class = "selected")
    current_interface("interactive")
  })
  
  observeEvent(input$syntax_choice, {
    removeClass(selector = "#interactive_choice", class = "selected")
    addClass(id = "syntax_choice", class = "selected")
    current_interface("syntax")
  })
  
  # Confirm interface selection
  observeEvent(input$confirm_interface, {
    if (current_interface() == "") {
      showNotification("Please select an interface first!", type = "warning")
      return()
    }
    interface_selected(TRUE)
    showNotification(paste("Switched to", current_interface(), "interface"), type = "message")
  })
  
  # Switch interface
  observeEvent(input$switch_interface, {
    interface_selected(FALSE)
    current_interface("")
    removeClass(selector = ".interface-choice-card", class = "selected")
  })
  
  # Reactive value to track sidebar state
  sidebar_collapsed <- reactiveVal(FALSE)
  
  # Toggle sidebar and auto-scroll when on Diagram tab
  observeEvent(input$main_tabs, {
    if (input$main_tabs == "Diagram") {
      sidebar_collapsed(TRUE)
      shinyjs::hide("sidebar", anim = TRUE, animType = "fade", time = 0.3)
      shinyjs::runjs("
        $('#main_panel').removeClass('col-sm-8').addClass('col-sm-12');
        $('#diagram_controls').show();
      ")
    } else {
      sidebar_collapsed(FALSE)
      shinyjs::show("sidebar", anim = TRUE, animType = "fade", time = 0.3)
      shinyjs::runjs("
        $('#main_panel').removeClass('col-sm-12').addClass('col-sm-8');
      ")
    }
  })
  
  # Toggle sidebar manually
  observeEvent(input$toggle_sidebar, {
    sidebar_collapsed(!sidebar_collapsed())
    if (sidebar_collapsed()) {
      shinyjs::hide("sidebar", anim = TRUE, animType = "fade", time = 0.3)
      shinyjs::runjs("
        $('#main_panel').removeClass('col-sm-8').addClass('col-sm-12');
      ")
    } else {
      shinyjs::show("sidebar", anim = TRUE, animType = "fade", time = 0.3)
      shinyjs::runjs("
        $('#main_panel').removeClass('col-sm-12').addClass('col-sm-8');
      ")
    }
  })
  
  # Read data with support for larger files and handle labelled data properly
  data <- reactive({
    req(input$datafile)
    ext <- tools::file_ext(input$datafile$name)
    
    tryCatch({
      if (ext == "csv") {
        read_csv(input$datafile$datapath, guess_max = 10000)
      } else if (ext %in% c("xlsx", "xls")) {
        read_excel(input$datafile$datapath, guess_max = 10000)
      } else if (ext %in% c("dta")) {
        read_dta(input$datafile$datapath)
      } else if (ext %in% c("sav", "zsav", "por")) {
        # Handle SPSS data with labelled variables
        df <- read_sav(input$datafile$datapath)
        # Convert labelled variables to factors or numeric for proper display
        df <- haven::zap_labels(df)  # Remove label attributes to avoid DataTable issues
        df
      } else {
        stop("Unsupported file format")
      }
    }, error = function(e) {
      showNotification(paste("Error reading file:", e$message), type = "error", duration = 10)
      return(NULL)
    })
  })
  
  # Variable selection UI for interactive interface
  output$variable_selection_ui <- renderUI({
    req(data())
    df <- data()
    var_names <- names(df)
    
    tagList(
      # Simple Mediation Variables
      conditionalPanel(
        condition = "input.analysis_type_interactive == 'Simple Mediation'",
        div(class = "variable-selection-box", 
            id = "x_box",
            onclick = "Shiny.setInputValue('select_variable', 'X', {priority: 'event'});",
            div(class = "variable-label", "Independent Variable (X)"),
            if(!is.null(selected_variables$X)) {
              div(class = "variable-name", selected_variables$X)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        ),
        div(class = "variable-selection-box", 
            id = "m_box",
            onclick = "Shiny.setInputValue('select_variable', 'M', {priority: 'event'});",
            div(class = "variable-label", "Mediator Variable (M)"),
            if(!is.null(selected_variables$M)) {
              div(class = "variable-name", selected_variables$M)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        ),
        div(class = "variable-selection-box", 
            id = "y_box",
            onclick = "Shiny.setInputValue('select_variable', 'Y', {priority: 'event'});",
            div(class = "variable-label", "Dependent Variable (Y)"),
            if(!is.null(selected_variables$Y)) {
              div(class = "variable-name", selected_variables$Y)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        )
      ),
      
      # Serial Mediation Variables
      conditionalPanel(
        condition = "input.analysis_type_interactive == 'Serial Mediation'",
        div(class = "variable-selection-box", 
            onclick = "Shiny.setInputValue('select_variable', 'X', {priority: 'event'});",
            div(class = "variable-label", "Independent Variable (X)"),
            if(!is.null(selected_variables$X)) {
              div(class = "variable-name", selected_variables$X)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        ),
        div(class = "variable-selection-box", 
            onclick = "Shiny.setInputValue('select_variable', 'M1', {priority: 'event'});",
            div(class = "variable-label", "Mediator 1 (M1)"),
            if(!is.null(selected_variables$M1)) {
              div(class = "variable-name", selected_variables$M1)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        ),
        div(class = "variable-selection-box", 
            onclick = "Shiny.setInputValue('select_variable', 'M2', {priority: 'event'});",
            div(class = "variable-label", "Mediator 2 (M2)"),
            if(!is.null(selected_variables$M2)) {
              div(class = "variable-name", selected_variables$M2)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        ),
        div(class = "variable-selection-box", 
            onclick = "Shiny.setInputValue('select_variable', 'Y', {priority: 'event'});",
            div(class = "variable-label", "Dependent Variable (Y)"),
            if(!is.null(selected_variables$Y)) {
              div(class = "variable-name", selected_variables$Y)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        )
      ),
      
      # Moderation Variables
      conditionalPanel(
        condition = "input.analysis_type_interactive == 'Moderation'",
        div(class = "variable-selection-box", 
            onclick = "Shiny.setInputValue('select_variable', 'X', {priority: 'event'});",
            div(class = "variable-label", "Independent Variable (X)"),
            if(!is.null(selected_variables$X)) {
              div(class = "variable-name", selected_variables$X)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        ),
        div(class = "variable-selection-box", 
            onclick = "Shiny.setInputValue('select_variable', 'W', {priority: 'event'});",
            div(class = "variable-label", "Moderator Variable (W)"),
            if(!is.null(selected_variables$W)) {
              div(class = "variable-name", selected_variables$W)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        ),
        div(class = "variable-selection-box", 
            onclick = "Shiny.setInputValue('select_variable', 'Y', {priority: 'event'});",
            div(class = "variable-label", "Dependent Variable (Y)"),
            if(!is.null(selected_variables$Y)) {
              div(class = "variable-name", selected_variables$Y)
            } else {
              div(class = "click-instruction", "Click to select variable")
            }
        )
      ),
      
      # Variable selection modal trigger
      actionButton("show_var_modal", "Select Variables", 
                   class = "btn-info btn-sm", icon = icon("list"))
    )
  })
  
  # Handle variable selection
  observeEvent(input$select_variable, {
    req(data())
    showModal(
      modalDialog(
        title = paste("Select", input$select_variable, "Variable"),
        selectInput("variable_choice", "Choose variable:", 
                    choices = names(data())),
        footer = tagList(
          modalButton("Cancel"),
          actionButton("confirm_variable", "Select", class = "btn-primary")
        )
      )
    )
  })
  
  observeEvent(input$confirm_variable, {
    var_type <- input$select_variable
    var_name <- input$variable_choice
    
    # Store the variable name
    selected_variables[[var_type]] <- var_name
    
    # Update the UI to show selected variable
    removeModal()
    
    # Update the box styling
    runjs(paste0("$('#", tolower(var_type), "_box').addClass('has-variable');"))
    
    # Update the box text to show selected variable
    runjs(paste0("
    $('#", tolower(var_type), "_box .variable-name').text('", var_name, "');
    $('#", tolower(var_type), "_box .click-instruction').hide();
  "))
    
    # Show notification
    showNotification(paste("Selected", var_name, "as", var_type), type = "message", duration = 2)
  })
  
  # Function to detect and convert categorical variables to dummy variables - IMPROVED VERSION
  process_covariates <- function(data, covariate_names) {
    if (is.null(covariate_names) || length(covariate_names) == 0) {
      return(list(data = data, covariate_terms = NULL))
    }
    
    # Create dummy variables for categorical covariates
    dummy_terms <- character()
    
    for (cov in covariate_names) {
      # Skip if covariate doesn't exist in data
      if (!cov %in% names(data)) {
        warning(paste("Covariate", cov, "not found in data"))
        next
      }
      
      var_data <- data[[cov]]
      
      # Check if variable is numeric (continuous)
      if (is.numeric(var_data)) {
        # Numeric variable - use as is
        dummy_terms <- c(dummy_terms, cov)
        next
      }
      
      # Handle categorical variables
      tryCatch({
        # Convert to factor for categorical variables
        if (inherits(var_data, "haven_labelled")) {
          var_data <- haven::as_factor(var_data)
        }
        
        if (is.character(var_data)) {
          var_data <- as.factor(var_data)
        }
        
        if (is.factor(var_data)) {
          # Get unique levels
          levels <- levels(var_data)
          
          # Only create dummies if there are multiple levels
          if (length(levels) > 1) {
            # Create dummy variables for each level except reference
            ref_level <- levels[1]
            
            for (i in 2:length(levels)) {
              # Create clean level name (remove special characters)
              level_clean <- gsub("[^A-Za-z0-9_]", "", as.character(levels[i]))
              if (nchar(level_clean) == 0) {
                level_clean <- paste0("level", i)
              }
              
              dummy_name <- paste0(cov, "_", level_clean)
              
              # Create dummy variable
              data[[dummy_name]] <- as.numeric(as.character(var_data) == levels[i])
              
              # Add to terms
              dummy_terms <- c(dummy_terms, dummy_name)
            }
          } else {
            # Single level - treat as constant (add but warn)
            warning(paste("Covariate", cov, "has only one level, treating as constant"))
            dummy_terms <- c(dummy_terms, cov)
          }
        } else {
          # Fallback: treat as numeric if conversion fails
          if (is.numeric(as.numeric(as.character(var_data)))) {
            data[[cov]] <- as.numeric(as.character(var_data))
            dummy_terms <- c(dummy_terms, cov)
          } else {
            warning(paste("Cannot process covariate", cov, "- skipping"))
          }
        }
      }, error = function(e) {
        warning(paste("Error processing covariate", cov, ":", e$message))
      })
    }
    
    # Remove any duplicate terms
    dummy_terms <- unique(dummy_terms)
    
    return(list(data = data, covariate_terms = dummy_terms))
  }
  
  # Generate model syntax from interactive selection - FIXED VERSION
  generate_model_syntax <- reactive({
    req(current_interface() == "interactive")
    
    # Get the actual variable names
    if (input$analysis_type_interactive == "Simple Mediation") {
      req(selected_variables$X, selected_variables$M, selected_variables$Y)
      
      # Get the actual variable names
      x_var <- selected_variables$X
      m_var <- selected_variables$M
      y_var <- selected_variables$Y
      
      syntax <- paste0(
        "# Simple Mediation Model\n",
        m_var, " ~ a*", x_var, "\n",
        y_var, " ~ b*", m_var, " + c*", x_var, "\n\n",
        "# User-friendly Indirect Effects\n",
        "Indirect_Effect_M := a*b\n",
        "Total_Effect := c + (a*b)"
      )
      
    } else if (input$analysis_type_interactive == "Serial Mediation") {
      req(selected_variables$X, selected_variables$M1, selected_variables$M2, selected_variables$Y)
      
      # Get the actual variable names
      x_var <- selected_variables$X
      m1_var <- selected_variables$M1
      m2_var <- selected_variables$M2
      y_var <- selected_variables$Y
      
      syntax <- paste0(
        "# Serial Mediation Model (", x_var, " -> ", m1_var, " -> ", m2_var, " -> ", y_var, ")\n",
        m1_var, " ~ a1*", x_var, "\n",
        m2_var, " ~ a2*", m1_var, " + d*", x_var, "\n",
        y_var, " ~ b1*", m1_var, " + b2*", m2_var, " + c*", x_var, "\n\n",
        "# User-friendly Indirect Effects\n",
        "Indirect_Effect_M1 := a1 * b1\n",
        "Indirect_Effect_M2 := d * b2\n", 
        "Indirect_Effect_Serial := a1 * a2 * b2\n",
        "Total_Indirect_Effect := Indirect_Effect_M1 + Indirect_Effect_M2 + Indirect_Effect_Serial\n",
        "Total_Effect := c + Total_Indirect_Effect"
      )
      
    } else if (input$analysis_type_interactive == "Moderation") {
      req(selected_variables$X, selected_variables$W, selected_variables$Y)
      
      # Get the actual variable names
      x_var <- selected_variables$X
      w_var <- selected_variables$W
      y_var <- selected_variables$Y
      
      # Clean variable names
      x_clean <- gsub("[^a-zA-Z0-9_]", "", x_var)
      w_clean <- gsub("[^a-zA-Z0-9_]", "", w_var)
      y_clean <- gsub("[^a-zA-Z0-9_]", "", y_var)
      
      # Create the interaction term name
      interaction_name <- paste0(x_clean, "_", w_clean)
      
      syntax <- paste0(
        "# Moderation Model with Automatic Interaction Term\n",
        "# The app automatically creates centered variables and interaction term\n\n",
        "# Main effects and interaction\n",
        y_clean, " ~ b1*", x_clean, " + b2*", w_clean, " + b3*", interaction_name, "\n\n",
        "# Simple slopes analysis (moderator centered)\n",
        "Simple_Slope_Low := b1 + b3*(-1)\n",
        "Simple_Slope_Avg := b1 + b3*(0)\n",
        "Simple_Slope_High := b1 + b3*(1)"
      )
    }
    
    return(syntax)
  })
  
  # Update model syntax when interactive variables change
  observe({
    if (current_interface() == "interactive") {
      syntax <- tryCatch({
        generate_model_syntax()
      }, error = function(e) {
        # Return empty syntax if variables not selected
        return("")
      })
      
      if (syntax != "") {
        updateTextAreaInput(session, "model_syntax", value = syntax)
      }
    }
  })
  
  
  # Interactive guide
  output$interactive_guide <- renderUI({
    guide_text <- switch(input$analysis_type_interactive,
                         "Simple Mediation" = tags$ul(
                           tags$li("Click on each box to select your variables"),
                           tags$li("X = Independent Variable (cause)"),
                           tags$li("M = Mediator Variable (mechanism)"),
                           tags$li("Y = Dependent Variable (outcome)"),
                           tags$li("The app will automatically generate the appropriate lavaan syntax")
                         ),
                         "Serial Mediation" = tags$ul(
                           tags$li("Click on each box to select your variables"),
                           tags$li("X = Independent Variable"),
                           tags$li("M1 = First Mediator"),
                           tags$li("M2 = Second Mediator"),
                           tags$li("Y = Dependent Variable"),
                           tags$li("Mediators should be in sequential order")
                         ),
                         "Moderation" = tags$ul(
                           tags$li("Click on each box to select your variables"),
                           tags$li("X = Independent Variable"),
                           tags$li("W = Moderator Variable"),
                           tags$li("Y = Dependent Variable"),
                           tags$li("Note: The app creates the interaction term in your data when you run the analysis")
                         )
    )
    return(guide_text)
  })
  
  # Variable preview with instruction
  output$var_preview <- renderUI({
    req(data())
    df <- data()
    var_names <- names(df)
    if(ncol(df) > 500) {
      showNotification("Warning: Dataset has more than 500 variables. Only the first 500 will be used.", 
                       type = "warning", duration = 10)
      var_names <- var_names[1:500]
    }
    
    if(length(var_names) > 15) {
      var_text <- paste0(paste(var_names[1:15], collapse = ", "), ", ... (+", length(var_names)-15, " more)")
    } else {
      var_text <- paste(var_names, collapse = ", ")
    }
    
    HTML(paste0("<strong>Variables in dataset (", ncol(df), "):</strong><br>", var_text,
                "<br><strong>Observations:</strong> ", nrow(df),
                "<br><br><em>Go to 'Variables in Data' tab to see all variable names.</em>"))
  })
  
  # Render DataTable with proper handling for labelled data
  output$var_table <- renderDT({
    req(data())
    df <- data()
    if(ncol(df) > 500) {
      df <- df[, 1:500]
    }
    # Convert any remaining labelled columns to factors for display
    df <- haven::zap_labels(df)
    datatable(df, 
              options = list(
                scrollX = TRUE,
                pageLength = 10,
                lengthMenu = c(5, 10, 15, 20)
              ))
  })
  
  output$data_summary <- renderPrint({
    req(data())
    df <- data()
    if(ncol(df) > 500) {
      df <- df[, 1:500]
    }
    
    cat("DATA SUMMARY\n")
    cat("------------\n")
    cat("Number of observations:", nrow(df), "\n")
    cat("Number of variables:", ncol(df), "\n\n")
    
    cat("VARIABLE TYPES:\n")
    var_types <- sapply(df, function(x) class(x)[1])
    print(table(var_types))
    
    cat("\nMISSING VALUES:\n")
    missing_count <- sapply(df, function(x) sum(is.na(x)))
    if(any(missing_count > 0)) {
      print(missing_count[missing_count > 0])
      cat("\nNote: The analysis will automatically remove rows with missing values.\n")
    } else {
      cat("No missing values detected.\n")
    }
  })
  
  # Render UI for covariate selection - allows both categorical and numerical
  output$covariate_ui <- renderUI({
    req(data())
    df <- data()
    var_names <- names(df)
    
    potential_keys <- c("X", "Y", "M", "M1", "M2", "W", "XW", "X_W", "interaction")
    covariate_choices <- setdiff(var_names, potential_keys)
    
    selectizeInput("covariates", "Select Covariates to Adjust For:",
                   choices = covariate_choices,
                   multiple = TRUE,
                   options = list(placeholder = 'Select variables to adjust for (categorical or numerical)'))
  })
  
  # Syntax examples - FIXED VERSION with user-friendly names
  output$simple_mediation_syntax <- renderText({
    "# Replace X, M, Y with your actual variable names
M ~ a*X
Y ~ b*M + c*X

# Define effects
Indirect_Effect_M := a*b
Total_Effect := c + (a*b)"
  })
  
  output$serial_mediation_syntax <- renderText({
    "# Replace X, M1, M2, Y with your actual variable names
M1 ~ a1*X
M2 ~ a2*M1 + d*X
Y ~ b1*M1 + b2*M2 + c*X

# Define user-friendly indirect effects
Indirect_Effect_M1 := a1 * b1
Indirect_Effect_M2 := d * b2
Indirect_Effect_Serial := a1 * a2 * b2
Total_Indirect_Effect := Indirect_Effect_M1 + Indirect_Effect_M2 + Indirect_Effect_Serial
Total_Effect := c + Total_Indirect_Effect"
  })
  
  output$moderation_syntax <- renderText({
    "# Replace X, W, XW, Y with your actual variable names
# Note: Create interaction term XW in your data first
Y ~ b1*X + b2*W + b3*XW

# Define simple slopes (assuming W is centered)
Simple_Slope_Low := b1 + b3*(-1)
Simple_Slope_Avg := b1 + b3*(0)
Simple_Slope_High := b1 + b3*(1)"
  })
  
  # How-to guide outputs - FIXED VERSION
  output$simple_mediation_howto1 <- renderText({
    "# Basic mediation syntax
M ~ a*X
Y ~ b*M + c*X"
  })
  
  output$simple_mediation_example <- renderText({
    "# Example with real variables
Anxiety ~ a*Stress
Performance ~ b*Anxiety + c*Stress"
  })
  
  output$simple_mediation_indirect <- renderText({
    "# Add these effect definitions
Indirect_Effect_M := a*b
Total_Effect := c + Indirect_Effect_M"
  })
  
  output$serial_mediation_howto1 <- renderText({
    "# Basic serial mediation syntax
M1 ~ a1*X
M2 ~ a2*M1 + d*X
Y ~ b1*M1 + b2*M2 + c*X"
  })
  
  output$serial_mediation_example <- renderText({
    "# Example with real variables
Motivation ~ a1*Calling
Engagement ~ a2*Motivation + d*Calling
Satisfaction ~ b1*Motivation + b2*Engagement + c*Calling"
  })
  
  output$serial_mediation_indirect <- renderText({
    "# Add these user-friendly effect definitions
Indirect_Effect_M1 := a1 * b1
Indirect_Effect_M2 := d * b2
Indirect_Effect_Serial := a1 * a2 * b2
Total_Indirect_Effect := Indirect_Effect_M1 + Indirect_Effect_M2 + Indirect_Effect_Serial
Total_Effect := c + Total_Indirect_Effect"
  })
  
  output$moderation_howto1 <- renderText({
    "# Basic moderation syntax
Y ~ b1*X + b2*W + b3*XW"
  })
  
  output$moderation_example <- renderText({
    "# Example with real variables
Performance ~ b1*Stress + b2*Support + b3*Stress_Support"
  })
  
  output$moderation_slopes <- renderText({
    "# Add these simple slopes
Simple_Slope_Low := b1 + b3*(-1)
Simple_Slope_Avg := b1 + b3*(0)
Simple_Slope_High := b1 + b3*(1)"
  })
  
  # Function to extract all variables from model syntax
  extract_all_variables <- function(model_syntax, data_vars) {
    equations <- strsplit(model_syntax, "\n")[[1]]
    equations <- equations[!grepl("^\\s*#", equations)]
    equations <- equations[nchar(trimws(equations)) > 0]
    
    all_vars <- character()
    lavaan_keywords <- c("a", "b", "c", "cp", "d", "a1", "a2", "b1", "b2", "b3", 
                         "indirect", "total", "std", "lowW", "avgW", "highW",
                         "indirect1", "indirect2", "indirect3", "total_indirect", "total_effect",
                         "Indirect_Effect_M", "Indirect_Effect_M1", "Indirect_Effect_M2", 
                         "Indirect_Effect_Serial", "Total_Indirect_Effect", "Total_Effect",
                         "Simple_Slope_Low", "Simple_Slope_Avg", "Simple_Slope_High")
    
    for(eq in equations) {
      clean_eq <- gsub("\\b[a-zA-Z][a-zA-Z0-9_]*\\*", "", eq)
      clean_eq <- gsub(":=", " ", clean_eq)
      clean_eq <- gsub("~", " ", clean_eq)
      clean_eq <- gsub("\\+", " ", clean_eq)
      
      vars_in_eq <- strsplit(clean_eq, "[^a-zA-Z0-9_.]")[[1]]
      vars_in_eq <- vars_in_eq[nchar(vars_in_eq) > 0]
      vars_in_eq <- vars_in_eq[!vars_in_eq %in% lavaan_keywords]
      
      all_vars <- c(all_vars, vars_in_eq)
    }
    
    unique_vars <- unique(all_vars)
    unique_vars <- unique_vars[unique_vars %in% data_vars]
    
    return(unique_vars)
  }
  
  # Function to compute comprehensive regression results
  compute_regression_results <- function() {
    req(model_fit(), data())
    fit <- model_fit()
    df <- data()
    
    if(is.null(fit)) return(NULL)
    
    tryCatch({
      model_syntax <- input$model_syntax
      data_vars <- names(df)
      all_vars <- extract_all_variables(model_syntax, data_vars)
      
      if(length(all_vars) < 2) {
        return(NULL)
      }
      
      df_clean <- na.omit(df[, all_vars])
      
      if(nrow(df_clean) < 10) {
        showNotification("Warning: Too few observations for regression analysis", type = "warning")
        return(NULL)
      }
      
      regression_results <- list()
      
      for(i in 1:length(all_vars)) {
        for(j in 1:length(all_vars)) {
          if(i != j) {
            dep_var <- all_vars[i]
            indep_var <- all_vars[j]
            
            formula <- as.formula(paste(dep_var, "~", indep_var))
            lm_fit <- lm(formula, data = df_clean)
            lm_summary <- summary(lm_fit)
            coef_table <- coef(lm_summary)
            
            if(indep_var %in% rownames(coef_table)) {
              x_coef <- coef_table[indep_var, ]
              
              sd_x <- sd(df_clean[[indep_var]], na.rm = TRUE)
              sd_y <- sd(df_clean[[dep_var]], na.rm = TRUE)
              std_beta <- x_coef["Estimate"] * (sd_x / sd_y)
              
              std_se <- x_coef["Std. Error"] * (sd_x / sd_y)
              std_t_value <- std_beta / std_se
              std_p_value <- 2 * pt(abs(std_t_value), df = lm_fit$df.residual, lower.tail = FALSE)
              
              ci_lower <- x_coef["Estimate"] - 1.96 * x_coef["Std. Error"]
              ci_upper <- x_coef["Estimate"] + 1.96 * x_coef["Std. Error"]
              std_ci_lower <- std_beta - 1.96 * std_se
              std_ci_upper <- std_beta + 1.96 * std_se
              
              result_key <- paste(dep_var, "~", indep_var)
              regression_results[[result_key]] <- list(
                dependent = dep_var,
                independent = indep_var,
                formula = paste(dep_var, "~", indep_var),
                unstandardized = list(
                  estimate = x_coef["Estimate"],
                  se = x_coef["Std. Error"],
                  t_value = x_coef["t value"],
                  p_value = x_coef["Pr(>|t|)"],
                  ci_lower = ci_lower,
                  ci_upper = ci_upper
                ),
                standardized = list(
                  estimate = std_beta,
                  se = std_se,
                  t_value = std_t_value,
                  p_value = std_p_value,
                  ci_lower = std_ci_lower,
                  ci_upper = std_ci_upper
                ),
                r_squared = lm_summary$r.squared,
                adj_r_squared = lm_summary$adj.r.squared,
                n_obs = nrow(df_clean)
              )
            }
          }
        }
      }
      
      return(regression_results)
      
    }, error = function(e) {
      showNotification(paste("Error computing regression results:", e$message), type = "warning")
      return(NULL)
    })
  }
  
  # Reactive for regression results
  regression_results <- reactive({
    req(model_fit())
    compute_regression_results()
  })
  
  # Enhanced model estimation with automatic interaction term creation for moderation
  model_fit <- eventReactive(input$run, {
    req(input$model_syntax, data())
    
    # For interactive interface, check if variables are selected
    if (current_interface() == "interactive") {
      # Debug output
      cat("Current interface:", current_interface(), "\n")
      cat("Analysis type:", input$analysis_type_interactive, "\n")
      cat("Selected X:", selected_variables$X, "\n")
      cat("Selected M:", selected_variables$M, "\n")
      cat("Selected Y:", selected_variables$Y, "\n")
      cat("Selected M1:", selected_variables$M1, "\n")
      cat("Selected M2:", selected_variables$M2, "\n")
      cat("Selected W:", selected_variables$W, "\n")
      
      if (input$analysis_type_interactive == "Simple Mediation") {
        if (is.null(selected_variables$X) || is.null(selected_variables$M) || is.null(selected_variables$Y)) {
          showNotification("Please select all required variables for the analysis (X, M, Y)", type = "error")
          return(NULL)
        }
      } else if (input$analysis_type_interactive == "Serial Mediation") {
        if (is.null(selected_variables$X) || is.null(selected_variables$M1) || 
            is.null(selected_variables$M2) || is.null(selected_variables$Y)) {
          showNotification("Please select all required variables for the analysis (X, M1, M2, Y)", type = "error")
          return(NULL)
        }
      } else if (input$analysis_type_interactive == "Moderation") {
        if (is.null(selected_variables$X) || is.null(selected_variables$W) || is.null(selected_variables$Y)) {
          showNotification("Please select all required variables for the analysis (X, W, Y)", type = "error")
          return(NULL)
        }
      }
    }
    
    showNotification("Running analysis... This may take a few moments.", 
                     type = "message", duration = NULL, id = "analysis_progress")
    
    df <- data()
    
    if(is.null(df) || nrow(df) == 0) {
      removeNotification("analysis_progress")
      showNotification("Error: No data loaded or data is empty", type = "error", duration = 10)
      return(NULL)
    }
    
    # For moderation analysis, automatically create interaction term
    if (current_interface() == "interactive" && input$analysis_type_interactive == "Moderation") {
      x_var <- selected_variables$X
      w_var <- selected_variables$W
      
      if (!is.null(x_var) && !is.null(w_var)) {
        # Clean variable names for interaction
        x_var_clean <- gsub("[^a-zA-Z0-9_]", "", x_var)
        w_var_clean <- gsub("[^a-zA-Z0-9_]", "", w_var)
        interaction_name <- paste0(x_var_clean, "_", w_var_clean)
        
        # Check if interaction term already exists
        if (!interaction_name %in% names(df)) {
          # Create centered variables for better interpretation
          x_centered <- df[[x_var]] - mean(df[[x_var]], na.rm = TRUE)
          w_centered <- df[[w_var]] - mean(df[[w_var]], na.rm = TRUE)
          
          # Create interaction term
          df[[interaction_name]] <- x_centered * w_centered
          
          showNotification(
            paste("✓ Automatically created interaction term:", interaction_name, 
                  "(centered variables for better interpretation)"),
            type = "message",
            duration = 5
          )
          
          # Add to selected variables for reference
          selected_variables$interaction <- interaction_name
        } else {
          showNotification(
            paste("Using existing interaction term:", interaction_name),
            type = "message",
            duration = 3
          )
          selected_variables$interaction <- interaction_name
        }
      }
    }
    
    # Process covariates (categorical and numerical) to create dummy variables
    processed_data <- df
    covariate_terms <- NULL
    
    if(!is.null(input$covariates) && length(input$covariates) > 0) {
      # First, ensure we only process valid covariates
      valid_covariates <- input$covariates[input$covariates %in% names(df)]
      
      if(length(valid_covariates) > 0) {
        proc_result <- process_covariates(df, valid_covariates)
        processed_data <- proc_result$data
        covariate_terms <- proc_result$covariate_terms
        
        # Show notification about processed covariates
        if(length(valid_covariates) != length(input$covariates)) {
          showNotification(
            paste("Note: Some covariates were not found in data:", 
                  paste(setdiff(input$covariates, valid_covariates), collapse=", ")),
            type = "warning",
            duration = 5
          )
        }
        
        if(length(covariate_terms) > 0) {
          showNotification(
            paste("Added", length(covariate_terms), "covariate term(s) to the model"),
            type = "message",
            duration = 3
          )
        }
      } else {
        showNotification("No valid covariates selected", type = "warning", duration = 3)
      }
    }
    
    # Remove missing values
    processed_data <- na.omit(processed_data)
    
    # Check if we have enough data after removing missing values
    if(nrow(processed_data) < 10) {
      removeNotification("analysis_progress")
      showNotification("Error: Too few observations after removing missing values", type = "error", duration = 10)
      return(NULL)
    }
    
    if(nrow(processed_data) < 10) {
      removeNotification("analysis_progress")
      showNotification("Error: Too few observations after removing missing values", type = "error", duration = 10)
      return(NULL)
    }
    
    if(nrow(processed_data) > 10000) {
      showNotification("Warning: Dataset has more than 10,000 observations. Only the first 10,000 will be used.", 
                       type = "warning", duration = 10)
      processed_data <- processed_data[1:10000, ]
    }
    
    if(ncol(processed_data) > 500) {
      processed_data <- processed_data[, 1:500]
    }
    
    # Enhanced syntax validation
    model_syntax <- input$model_syntax
    
    if(nchar(model_syntax) < 5) {
      removeNotification("analysis_progress")
      showNotification("Error: Model syntax is too short or empty", type = "error", duration = 10)
      return(NULL)
    }
    
    if(!grepl("[~=:]", model_syntax)) {
      removeNotification("analysis_progress")
      showNotification("Error: Model syntax should contain SEM operators (~, :=, or =)", type = "error", duration = 10)
      return(NULL)
    }
    
    # Add covariates if specified (using processed dummy variable names)
    if(!is.null(covariate_terms) && length(covariate_terms) > 0) {
      cov_str <- paste(covariate_terms, collapse = " + ")
      
      equations <- strsplit(model_syntax, "\n")[[1]]
      reg_equations <- grep("~", equations, value = TRUE)
      
      for(i in seq_along(reg_equations)) {
        eq <- reg_equations[i]
        if(!grepl(paste(covariate_terms, collapse = "|"), eq)) {
          new_eq <- paste0(eq, " + ", cov_str)
          model_syntax <- gsub(eq, new_eq, model_syntax, fixed = TRUE)
        }
      }
    }
    
    tryCatch({
      # Extract variables from syntax with better error handling
      equations <- strsplit(model_syntax, "\n")[[1]]
      equations <- equations[!grepl("^\\s*#", equations)]
      equations <- equations[nchar(trimws(equations)) > 0]
      
      all_vars <- character()
      lavaan_keywords <- c("a", "b", "c", "cp", "d", "a1", "a2", "b1", "b2", "b3", 
                           "indirect", "total", "std", "lowW", "avgW", "highW",
                           "indirect1", "indirect2", "indirect3", "total_indirect", "total_effect",
                           "Indirect_Effect_M", "Indirect_Effect_M1", "Indirect_Effect_M2", 
                           "Indirect_Effect_Serial", "Total_Indirect_Effect", "Total_Effect",
                           "Simple_Slope_Low", "Simple_Slope_Avg", "Simple_Slope_High")
      
      for(eq in equations) {
        # Remove operators and keywords
        clean_eq <- gsub("\\b[a-zA-Z][a-zA-Z0-9_]*\\*", "", eq)
        clean_eq <- gsub(":=", " ", clean_eq)
        clean_eq <- gsub("~", " ", clean_eq)
        clean_eq <- gsub("\\+", " ", clean_eq)
        clean_eq <- gsub("\\*", " ", clean_eq)
        
        # Split and clean
        vars_in_eq <- strsplit(clean_eq, "[^a-zA-Z0-9_.]")[[1]]
        vars_in_eq <- vars_in_eq[nchar(vars_in_eq) > 0]
        
        # Remove lavaan keywords and numeric values (like 1, 0)
        vars_in_eq <- vars_in_eq[!vars_in_eq %in% lavaan_keywords]
        vars_in_eq <- vars_in_eq[!grepl("^[0-9]+$", vars_in_eq)]  # Remove pure numbers
        
        all_vars <- c(all_vars, vars_in_eq)
      }
      
      data_vars <- unique(all_vars)
      available_vars <- names(processed_data)
      missing_vars <- setdiff(data_vars, available_vars)
      
      # Filter out any variables that are actually dummy variables we created
      if(length(missing_vars) > 0 && !is.null(covariate_terms)) {
        # Check if missing vars are from covariate dummies that weren't created
        missing_vars <- missing_vars[!missing_vars %in% covariate_terms]
      }
      
      if(length(missing_vars) > 0) {
        removeNotification("analysis_progress")
        
        # Provide helpful error message
        error_msg <- paste("Variables not found in data:", paste(missing_vars, collapse = ", "))
        
        # Add suggestions for common issues
        if(any(grepl("_", missing_vars))) {
          error_msg <- paste0(error_msg, "\n\nNote: Some variables appear to be interaction terms. For moderation analysis, the app automatically creates these when you select variables.")
        }
        
        showNotification(error_msg, type = "error", duration = 15)
        return(NULL)
      }
      
      fit_result <- NULL
      
      # Enhanced model estimation with better error handling
      if(input$use_bootstrap) {
        fit_result <- tryCatch({
          sem(model = model_syntax, 
              data = processed_data, 
              std.lv = TRUE,
              se = "bootstrap", 
              bootstrap = min(input$bootstrap_samples, 5000),
              test = "standard",
              estimator = "ML",
              verbose = FALSE,
              optim.method = "NLMINB")
        }, error = function(e) {
          showNotification("Bootstrap failed, trying without bootstrap...", type = "warning", duration = 5)
          sem(model = model_syntax, 
              data = processed_data, 
              std.lv = TRUE,
              estimator = "ML",
              verbose = FALSE)
        })
      } else {
        fit_result <- sem(model = model_syntax, 
                          data = processed_data, 
                          std.lv = TRUE,
                          estimator = "ML",
                          verbose = FALSE)
      }
      
      if(!inherits(fit_result, "lavaan") || !lavInspect(fit_result, "converged")) {
        removeNotification("analysis_progress")
        showNotification("Warning: Model did not converge properly or returned invalid object", type = "warning", duration = 10)
        return(NULL)
      }
      
      # Test if fit measures can be computed
      test_fit <- tryCatch({
        fitMeasures(fit_result)
        TRUE
      }, error = function(e) {
        FALSE
      })
      
      if(!test_fit) {
        removeNotification("analysis_progress")
        showNotification("Error: Model estimation produced invalid results", type = "error", duration = 10)
        return(NULL)
      }
      
      removeNotification("analysis_progress")
      showNotification("Analysis completed successfully!", type = "message", duration = 5)
      
      updateTabsetPanel(session, "main_tabs", selected = "Results")
      
      return(fit_result)
      
    }, error = function(e) {
      removeNotification("analysis_progress")
      showNotification(paste("Error in model estimation:", e$message), 
                       type = "error", duration = 15)
      return(NULL)
    }, warning = function(w) {
      showNotification(paste("Warning:", w$message), type = "warning", duration = 10)
    })
  })
  
  # Dynamic UI outputs based on output options
  output$fitUI <- renderUI({
    req(model_fit())
    if ("fit" %in% input$output_options) {
      div(
        h3("Model Fit Indices", icon("check-circle")),
        div(verbatimTextOutput("fitText"), class = "resizable-text")
      )
    }
  })
  
  output$mediationTableUI <- renderUI({
    req(model_fit())
    if ("paths" %in% input$output_options) {
      div(
        h3("Path Estimates", icon("table")),
        DTOutput("mediationTable")
      )
    }
  })
  
  output$regressionTableUI <- renderUI({
    req(model_fit())
    if ("paths" %in% input$output_options) {
      div(
        h3("Regression Results", icon("arrow-right")),
        DTOutput("regressionTable")
      )
    }
  })
  
  output$effectsUI <- renderUI({
    req(model_fit())
    if ("total" %in% input$output_options || "direct" %in% input$output_options || "indirect" %in% input$output_options) {
      div(
        h3("Effects Analysis", icon("project-diagram")),
        div(verbatimTextOutput("effectsText"), class = "resizable-text")
      )
    }
  })
  
  # Output functions with improved null checks and error handling
  output$fitText <- renderPrint({
    req(model_fit())
    fit <- model_fit()
    
    if(is.null(fit) || !inherits(fit, "lavaan")) {
      cat("No valid model results available. Please check your model syntax and data.\n")
      cat("Common issues:\n")
      cat("1. Variables in model syntax don't match data variable names\n")
      cat("2. Model syntax has errors\n")
      cat("3. Not enough observations\n")
      cat("4. Multicollinearity or other estimation problems\n")
      return()
    }
    
    tryCatch({
      fit_measures <- fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "rmsea.ci.lower", "rmsea.ci.upper", "srmr"))
      
      cat("MODEL FIT INDICES:\n")
      cat(sprintf("Chi-square(%.0f) = %.3f, p = %.3f\n", fit_measures["df"], fit_measures["chisq"], fit_measures["pvalue"]))
      cat(sprintf("CFI = %.3f, TLI = %.3f\n", fit_measures["cfi"], fit_measures["tli"]))
      cat(sprintf("RMSEA = %.3f, 90%% CI [%.3f, %.3f]\n", fit_measures["rmsea"], fit_measures["rmsea.ci.lower"], fit_measures["rmsea.ci.upper"]))
      cat(sprintf("SRMR = %.3f\n\n", fit_measures["srmr"]))
      
      cat("Fit Interpretation:\n")
      if (fit_measures["cfi"] > 0.90 && fit_measures["rmsea"] <= 0.08) {
        cat("Excellent model fit (CFI > 0.90, RMSEA <= 0.08)\n")
      } else if (fit_measures["cfi"] > 0.90 && fit_measures["rmsea"] < 0.08) {
        cat("Acceptable model fit (CFI > 0.90, RMSEA < 0.08)\n")
      } else {
        cat("Poor model fit - consider revising your model\n")
      }
      
      if(input$use_bootstrap) {
        cat(sprintf("\nBootstrap Results (based on %d samples):\n", input$bootstrap_samples))
        cat("Note: Bootstrap provides more robust standard errors and confidence intervals\n")
      } else {
        cat("\nNote: Results are based on model-based standard errors \n")
      }
      
      if(!is.null(input$covariates) && length(input$covariates) > 0) {
        cat("\nCovariates adjusted for in the analysis:\n")
        cat(paste(input$covariates, collapse = ", "), "\n")
      }
    }, error = function(e) {
      cat("Error calculating fit measures:", e$message, "\n")
      cat("The model object may be corrupted. Please try running the analysis again.\n")
    })
  })
  
  # Create path estimates table (with bivariate associations for moderation)
  output$mediationTable <- renderDT({
    req(model_fit())
    fit <- model_fit()
    
    if(is.null(fit)) {
      return(NULL)
    }
    
    tryCatch({
      if(input$use_bootstrap) {
        pe <- parameterEstimates(fit, 
                                 standardized = TRUE, 
                                 ci = TRUE, 
                                 boot.ci.type = "perc",
                                 level = 0.95)
      } else {
        pe <- parameterEstimates(fit, 
                                 standardized = TRUE, 
                                 ci = TRUE)
      }
      
      # For moderation analysis, replace main effects with bivariate associations
      if(current_interface() == "interactive" && input$analysis_type_interactive == "Moderation") {
        
        # Get the original path estimates
        mediation_paths <- pe[pe$op == "~", ]
        indirect_effects <- pe[pe$op == ":=" & grepl("Indirect_Effect|Total_Effect|Simple_Slope", pe$lhs), ]
        
        # Identify which rows are the main effects (X → Y and W → Y)
        x_var <- selected_variables$X
        w_var <- selected_variables$W
        y_var <- selected_variables$Y
        
        # Compute bivariate associations using simple linear regression
        df <- data()
        
        # Function to compute bivariate regression
        compute_bivariate <- function(dep_var, indep_var, data) {
          formula <- as.formula(paste(dep_var, "~", indep_var))
          lm_fit <- lm(formula, data = data, na.action = na.omit)
          summary_fit <- summary(lm_fit)
          
          if (indep_var %in% rownames(coef(summary_fit))) {
            coefs <- coef(summary_fit)[indep_var, ]
            est <- coefs["Estimate"]
            se <- coefs["Std. Error"]
            t_val <- coefs["t value"]
            p_val <- coefs["Pr(>|t|)"]
            
            # Calculate confidence intervals
            ci_lower <- est - 1.96 * se
            ci_upper <- est + 1.96 * se
            
            # Calculate standardized beta
            sd_x <- sd(data[[indep_var]], na.rm = TRUE)
            sd_y <- sd(data[[dep_var]], na.rm = TRUE)
            std_beta <- est * (sd_x / sd_y)
            
            # Calculate standardized SE
            std_se <- std_beta / t_val
            
            return(list(
              lhs = dep_var,
              rhs = indep_var,
              op = "~",
              est = est,
              std.all = std_beta,
              se = se,
              pvalue = p_val,
              ci.lower = ci_lower,
              ci.upper = ci_upper
            ))
          }
          return(NULL)
        }
        
        # Compute bivariate associations
        bivariate_xy <- compute_bivariate(y_var, x_var, df)
        bivariate_wy <- compute_bivariate(y_var, w_var, df)
        
        # Create a modified version of mediation_paths
        modified_paths <- mediation_paths
        
        # Replace X → Y and W → Y with bivariate results
        if(!is.null(bivariate_xy)) {
          xy_rows <- which(modified_paths$lhs == y_var & modified_paths$rhs == x_var)
          if(length(xy_rows) > 0) {
            for(i in xy_rows) {
              modified_paths$est[i] <- bivariate_xy$est
              modified_paths$std.all[i] <- bivariate_xy$std.all
              modified_paths$se[i] <- bivariate_xy$se
              modified_paths$pvalue[i] <- bivariate_xy$pvalue
              modified_paths$ci.lower[i] <- bivariate_xy$ci.lower
              modified_paths$ci.upper[i] <- bivariate_xy$ci.upper
            }
          } else {
            # If not found, add as new rows
            new_row <- data.frame(
              lhs = bivariate_xy$lhs,
              op = bivariate_xy$op,
              rhs = bivariate_xy$rhs,
              est = bivariate_xy$est,
              std.all = bivariate_xy$std.all,
              se = bivariate_xy$se,
              pvalue = bivariate_xy$pvalue,
              ci.lower = bivariate_xy$ci.lower,
              ci.upper = bivariate_xy$ci.upper,
              stringsAsFactors = FALSE
            )
            modified_paths <- rbind(modified_paths, new_row)
          }
        }
        
        if(!is.null(bivariate_wy)) {
          wy_rows <- which(modified_paths$lhs == y_var & modified_paths$rhs == w_var)
          if(length(wy_rows) > 0) {
            for(i in wy_rows) {
              modified_paths$est[i] <- bivariate_wy$est
              modified_paths$std.all[i] <- bivariate_wy$std.all
              modified_paths$se[i] <- bivariate_wy$se
              modified_paths$pvalue[i] <- bivariate_wy$pvalue
              modified_paths$ci.lower[i] <- bivariate_wy$ci.lower
              modified_paths$ci.upper[i] <- bivariate_wy$ci.upper
            }
          } else {
            # If not found, add as new rows
            new_row <- data.frame(
              lhs = bivariate_wy$lhs,
              op = bivariate_wy$op,
              rhs = bivariate_wy$rhs,
              est = bivariate_wy$est,
              std.all = bivariate_wy$std.all,
              se = bivariate_wy$se,
              pvalue = bivariate_wy$pvalue,
              ci.lower = bivariate_wy$ci.lower,
              ci.upper = bivariate_wy$ci.upper,
              stringsAsFactors = FALSE
            )
            modified_paths <- rbind(modified_paths, new_row)
          }
        }
        
        # Keep the interaction effect and simple slopes as is
        interaction_effect <- mediation_paths[grepl("_", mediation_paths$rhs) & 
                                                !mediation_paths$rhs %in% c(x_var, w_var), ]
        simple_slopes <- indirect_effects
        
        # Combine everything
        relevant_paths <- rbind(modified_paths, simple_slopes)
        
        # Remove duplicates (if any)
        relevant_paths <- unique(relevant_paths)
        
      } else {
        # For mediation analyses, keep original behavior
        mediation_paths <- pe[pe$op == "~", ]
        indirect_effects <- pe[pe$op == ":=" & grepl("Indirect_Effect|Total_Effect|Simple_Slope", pe$lhs), ]
        relevant_paths <- rbind(mediation_paths, indirect_effects)
      }
      
      # Create the pathway names
      relevant_paths$Pathway <- ifelse(relevant_paths$op == "~", 
                                       paste(relevant_paths$lhs, "<-", relevant_paths$rhs),
                                       paste(relevant_paths$lhs, ":=", relevant_paths$rhs))
      
      # Prepare table data
      table_data <- relevant_paths[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
      names(table_data) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
      
      table_data$`p-value` <- format.pval(table_data$`p-value`, digits = 3, eps = 0.001)
      table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")] <- 
        round(table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")], 3)
      
      # Add a note for moderation analysis
      if(current_interface() == "interactive" && input$analysis_type_interactive == "Moderation") {
        table_data <- rbind(
          table_data,
          data.frame(
            Pathway = "Note: X → Y and W → Y are bivariate associations from simple linear regression (not controlling for other variables)",
            Estimate = NA,
            `Std. Estimate` = NA,
            SE = NA,
            `p-value` = NA,
            `CI Lower` = NA,
            `CI Upper` = NA,
            check.names = FALSE
          )
        )
      }
      
      datatable(
        table_data,
        rownames = FALSE,
        extensions = 'Buttons',
        options = list(
          dom = 'Bfrtip',
          buttons = c('copy', 'csv', 'excel', 'pdf'),
          pageLength = 10,
          scrollX = TRUE,
          autoWidth = TRUE,
          columnDefs = list(
            list(width = '200px', targets = 0),
            list(className = 'dt-center', targets = 1:6)
          )
        ),
        class = 'stripe hover'
      ) %>% 
        formatStyle(
          'p-value',
          backgroundColor = styleInterval(
            c(0.001, 0.01, 0.05), 
            c('#FF6B6B', '#FFA3A3', '#FFD6A5', '#C8E7A5')
          )
        )
    }, error = function(e) {
      showNotification(paste("Error creating mediation table:", e$message), type = "error")
      return(NULL)
    })
  })
  
  # Create comprehensive regression table
  output$regressionTable <- renderDT({
    req(regression_results())
    results <- regression_results()
    
    if(is.null(results) || length(results) == 0) {
      return(NULL)
    }
    
    table_data <- data.frame()
    
    for(key in names(results)) {
      res <- results[[key]]
      
      unstd_row <- data.frame(
        Regression = res$formula,
        Coefficient_Type = "Unstandardized",
        Estimate = round(res$unstandardized$estimate, 3),
        Std_Error = round(res$unstandardized$se, 3),
        CI_Lower = round(res$unstandardized$ci_lower, 3),
        CI_Upper = round(res$unstandardized$ci_upper, 3),
        p_Value = format.pval(res$unstandardized$p_value, digits = 3, eps = 0.001),
        R_Squared = round(res$r_squared, 3),
        N = res$n_obs,
        stringsAsFactors = FALSE
      )
      
      std_row <- data.frame(
        Regression = res$formula,
        Coefficient_Type = "Standardized",
        Estimate = round(res$standardized$estimate, 3),
        Std_Error = round(res$standardized$se, 3),
        CI_Lower = round(res$standardized$ci_lower, 3),
        CI_Upper = round(res$standardized$ci_upper, 3),
        p_Value = format.pval(res$standardized$p_value, digits = 3, eps = 0.001),
        R_Squared = round(res$r_squared, 3),
        N = res$n_obs,
        stringsAsFactors = FALSE
      )
      
      table_data <- rbind(table_data, unstd_row, std_row)
    }
    
    names(table_data) <- c("Regression", "Coefficient Type", "Estimate", "SE", "CI Lower", "CI Upper", 
                           "p-value", "R²", "N")
    
    datatable(
      table_data,
      rownames = FALSE,
      extensions = 'Buttons',
      options = list(
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel', 'pdf'),
        pageLength = 20,
        scrollX = TRUE,
        autoWidth = TRUE,
        columnDefs = list(
          list(width = '250px', targets = 0),
          list(width = '150px', targets = 1),
          list(className = 'dt-center', targets = 2:8)
        )
      ),
      class = 'stripe hover'
    ) %>% 
      formatStyle(
        'p-value',
        backgroundColor = styleInterval(
          c(0.001, 0.01, 0.05), 
          c('#FF6B6B', '#FFA3A3', '#FFD6A5', '#C8E7A5')
        )
      )
  })
  
  output$effectsText <- renderPrint({
    req(model_fit())
    fit <- model_fit()
    
    if(is.null(fit)) {
      cat("No model results available.\n")
      return()
    }
    
    tryCatch({
      if(input$use_bootstrap) {
        pe <- parameterEstimates(fit, 
                                 standardized = TRUE, 
                                 ci = TRUE, 
                                 boot.ci.type = "perc",
                                 level = 0.95)
      } else {
        pe <- parameterEstimates(fit, 
                                 standardized = TRUE, 
                                 ci = TRUE)
      }
      
      # Filter effects based on output options
      defined <- pe[pe$op == ":=", ]
      effects_to_show <- data.frame()
      
      if ("total" %in% input$output_options) {
        total_effects <- defined[grepl("Total_Effect|total_effect", defined$lhs, ignore.case = TRUE), ]
        effects_to_show <- rbind(effects_to_show, total_effects)
      }
      
      if ("direct" %in% input$output_options) {
        direct_effects <- pe[pe$op == "~" & !grepl("indirect|Indirect", pe$lhs, ignore.case = TRUE), ]
        if(nrow(direct_effects) > 0) {
          cat("DIRECT EFFECTS:\n\n")
          print(direct_effects[, c("lhs", "op", "rhs", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")], 
                row.names = FALSE)
          cat("\n")
        }
      }
      
      if ("indirect" %in% input$output_options) {
        indirect_effects <- defined[grepl("Indirect_Effect|indirect|Serial", defined$lhs, ignore.case = TRUE), ]
        effects_to_show <- rbind(effects_to_show, indirect_effects)
      }
      
      if (nrow(effects_to_show) > 0) {
        cat("EFFECTS ANALYSIS:\n\n")
        print(effects_to_show[, c("lhs", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")], 
              row.names = FALSE)
      }
      
      # Show other defined effects if any
      other_effects <- defined[!grepl("Total_Effect|total_effect|Indirect_Effect|indirect|Serial", defined$lhs, ignore.case = TRUE), ]
      if (nrow(other_effects) > 0) {
        cat("\nOTHER DEFINED EFFECTS:\n\n")
        print(other_effects[, c("lhs", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")], 
              row.names = FALSE)
      }
      
      if(input$use_bootstrap) {
        cat("\nNote: Effects estimates use bootstrap standard errors and confidence intervals\n")
      } else {
        cat("\nNote: Effects estimates use model-based standard errors \n")
      }
    }, error = function(e) {
      cat("Error generating effects text:", e$message, "\n")
    })
  })
  
  
  # Function to compute simple slopes for plotting
  compute_simple_slopes_data <- function() {
    req(model_fit())
    fit <- model_fit()
    df <- data()
    
    # Get the parameter estimates
    if(input$use_bootstrap) {
      pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
    } else {
      pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
    }
    
    # Get variable names
    if (current_interface() == "interactive") {
      x_var <- selected_variables$X
      w_var <- selected_variables$W
      y_var <- selected_variables$Y
    } else {
      # Try to extract from syntax
      syntax <- input$model_syntax
      # Simple extraction - look for typical moderation patterns
      x_var <- "X"
      w_var <- "W"
      y_var <- "Y"
    }
    
    # Get the coefficients
    x_coef <- pe[pe$op == "~" & pe$rhs == x_var & pe$lhs == y_var, "est"]
    w_coef <- pe[pe$op == "~" & pe$rhs == w_var & pe$lhs == y_var, "est"]
    xw_coef <- pe[pe$op == "~" & grepl(paste0(x_var, "_", w_var), pe$rhs) & pe$lhs == y_var, "est"]
    
    if(length(x_coef) == 0 || length(w_coef) == 0 || length(xw_coef) == 0) {
      return(NULL)
    }
    
    # Get moderator values
    w_values <- df[[w_var]]
    w_mean <- mean(w_values, na.rm = TRUE)
    w_sd <- sd(w_values, na.rm = TRUE)
    
    # Create sequence of X values
    x_values <- df[[x_var]]
    x_min <- min(x_values, na.rm = TRUE)
    x_max <- max(x_values, na.rm = TRUE)
    x_seq <- seq(x_min, x_max, length.out = 100)
    
    # Compute slopes at different moderator levels
    w_low <- w_mean - w_sd
    w_avg <- w_mean
    w_high <- w_mean + w_sd
    
    # Calculate predicted Y values
    y_low <- x_coef * x_seq + w_coef * w_low + xw_coef * (x_seq * w_low)
    y_avg <- x_coef * x_seq + w_coef * w_avg + xw_coef * (x_seq * w_avg)
    y_high <- x_coef * x_seq + w_coef * w_high + xw_coef * (x_seq * w_high)
    
    # Create data frame for plotting
    plot_data <- data.frame(
      X = rep(x_seq, 3),
      Y = c(y_low, y_avg, y_high),
      Moderator = factor(
        rep(c("Low", "Average", "High"), each = length(x_seq)),
        levels = c("Low", "Average", "High")
      )
    )
    
    # Add original data points
    original_data <- data.frame(
      X = df[[x_var]],
      Y = df[[y_var]]
    )
    
    return(list(
      plot_data = plot_data,
      original_data = original_data,
      coefficients = list(
        b1 = as.numeric(x_coef),
        b2 = as.numeric(w_coef),
        b3 = as.numeric(xw_coef),
        w_mean = w_mean,
        w_sd = w_sd
      ),
      variable_names = list(
        X = x_var,
        W = w_var,
        Y = y_var
      )
    ))
  }
  
  # Render the slope plot
  output$slopePlot <- renderPlot({
    # Check if it's a moderation analysis
    if (!is_moderation()) {
      return(NULL)
    }
    
    req(model_fit())
    
    slope_data <- compute_simple_slopes_data()
    
    if(is.null(slope_data)) {
      return(NULL)
    }
    
    # Create the plot
    p <- ggplot(slope_data$plot_data, aes(x = X, y = Y, color = Moderator, group = Moderator)) +
      geom_line(size = input$slope_line_width) +
      geom_point(data = slope_data$original_data, 
                 aes(x = X, y = Y), 
                 color = "grey50", 
                 alpha = input$slope_alpha,
                 size = input$slope_point_size,
                 inherit.aes = FALSE) +
      scale_color_manual(
        values = c(
          "Low" = input$slope_low_color,
          "Average" = input$slope_avg_color,
          "High" = input$slope_high_color
        ),
        labels = c(
          "Low" = input$mod_low_label,
          "Average" = input$mod_avg_label,
          "High" = input$mod_high_label
        )
      ) +
      labs(
        title = "Simple Slopes Analysis",
        subtitle = paste("Moderator:", slope_data$variable_names$W),
        x = ifelse(nchar(input$x_axis_label) > 0, 
                   input$x_axis_label, 
                   slope_data$variable_names$X),
        y = ifelse(nchar(input$y_axis_label) > 0, 
                   input$y_axis_label, 
                   slope_data$variable_names$Y),
        color = "Moderator Level"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        plot.subtitle = element_text(size = 12, hjust = 0.5, color = "grey50"),
        legend.position = "bottom",
        legend.title = element_text(face = "bold"),
        panel.grid.minor = element_blank(),
        panel.border = element_rect(fill = NA, color = "grey80"),
        axis.title = element_text(size = 12, face = "bold"),
        axis.text = element_text(size = 10)
      ) +
      annotate(
        "text",
        x = Inf,
        y = -Inf,
        label = paste(
          "Interaction (X×W):", round(slope_data$coefficients$b3, 3),
          "\nLow: b1 + b3×(-1 SD) =", round(slope_data$coefficients$b1 + slope_data$coefficients$b3 * -1, 3),
          "\nAvg: b1 =", round(slope_data$coefficients$b1, 3),
          "\nHigh: b1 + b3×(+1 SD) =", round(slope_data$coefficients$b1 + slope_data$coefficients$b3 * 1, 3)
        ),
        hjust = 1,
        vjust = -0.5,
        size = 3,
        color = "grey30"
      )
    
    p
  }, height = function() {
    input$slope_plot_height
  })
  
  # Simple slopes values output
  output$simple_slopes_values <- renderPrint({
    if (!is_moderation()) {
      cat("Moderation analysis not selected. Please select 'Moderation' as your analysis type to view simple slopes values.")
      return()
    }
    
    req(model_fit())
    
    slope_data <- compute_simple_slopes_data()
    
    if(is.null(slope_data)) {
      cat("Unable to compute simple slopes. Please check your model.")
      return()
    }
    
    cat("SIMPLE SLOPES ANALYSIS\n")
    cat("======================\n\n")
    
    cat("Model: Y = b1*X + b2*W + b3*(X×W)\n\n")
    
    cat("Coefficients:\n")
    cat(sprintf("  b1 (X → Y): %.4f\n", slope_data$coefficients$b1))
    cat(sprintf("  b2 (W → Y): %.4f\n", slope_data$coefficients$b2))
    cat(sprintf("  b3 (X×W → Y): %.4f\n\n", slope_data$coefficients$b3))
    
    cat("Moderator Statistics:\n")
    cat(sprintf("  Mean of %s: %.4f\n", slope_data$variable_names$W, slope_data$coefficients$w_mean))
    cat(sprintf("  SD of %s: %.4f\n\n", slope_data$variable_names$W, slope_data$coefficients$w_sd))
    
    cat("Simple Slopes:\n")
    cat(sprintf("  At -1 SD (Low %s): slope = %.4f\n", 
                slope_data$variable_names$W, 
                slope_data$coefficients$b1 + slope_data$coefficients$b3 * -1))
    cat(sprintf("  At Mean (Average %s): slope = %.4f\n", 
                slope_data$variable_names$W, 
                slope_data$coefficients$b1))
    cat(sprintf("  At +1 SD (High %s): slope = %.4f\n\n", 
                slope_data$variable_names$W, 
                slope_data$coefficients$b1 + slope_data$coefficients$b3 * 1))
    
    cat("Interpretation:\n")
    if(slope_data$coefficients$b3 != 0) {
      if(slope_data$coefficients$b3 > 0) {
        cat("  Positive interaction: The effect of X on Y INCREASES as W increases.\n")
      } else {
        cat("  Negative interaction: The effect of X on Y DECREASES as W increases.\n")
      }
    } else {
      cat("  No significant interaction: The effect of X on Y is constant across levels of W.\n")
    }
  })
  
  # Download handler for slope plot
  output$download_slope_plot <- downloadHandler(
    filename = function() {
      paste0("Simple_Slopes_Plot_", Sys.Date(), ".png")
    },
    content = function(file) {
      if (!is_moderation()) {
        showNotification("Slope plot only available for moderation analysis", type = "error")
        return()
      }
      
      req(model_fit())
      
      slope_data <- compute_simple_slopes_data()
      
      if(is.null(slope_data)) {
        showNotification("Cannot generate slope plot - no valid moderation model", type = "error")
        return()
      }
      
      # Create plot with current settings
      p <- ggplot(slope_data$plot_data, aes(x = X, y = Y, color = Moderator, group = Moderator)) +
        geom_line(size = input$slope_line_width) +
        geom_point(data = slope_data$original_data, 
                   aes(x = X, y = Y), 
                   color = "grey50", 
                   alpha = input$slope_alpha,
                   size = input$slope_point_size,
                   inherit.aes = FALSE) +
        scale_color_manual(
          values = c(
            "Low" = input$slope_low_color,
            "Average" = input$slope_avg_color,
            "High" = input$slope_high_color
          ),
          labels = c(
            "Low" = input$mod_low_label,
            "Average" = input$mod_avg_label,
            "High" = input$mod_high_label
          )
        ) +
        labs(
          title = "Simple Slopes Analysis",
          subtitle = paste("Moderator:", slope_data$variable_names$W),
          x = ifelse(nchar(input$x_axis_label) > 0, 
                     input$x_axis_label, 
                     slope_data$variable_names$X),
          y = ifelse(nchar(input$y_axis_label) > 0, 
                     input$y_axis_label, 
                     slope_data$variable_names$Y),
          color = "Moderator Level"
        ) +
        theme_minimal() +
        theme(
          plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
          plot.subtitle = element_text(size = 12, hjust = 0.5, color = "grey50"),
          legend.position = "bottom",
          legend.title = element_text(face = "bold"),
          panel.grid.minor = element_blank(),
          panel.border = element_rect(fill = NA, color = "grey80"),
          axis.title = element_text(size = 12, face = "bold"),
          axis.text = element_text(size = 10)
        ) +
        annotate(
          "text",
          x = Inf,
          y = -Inf,
          label = paste(
            "Interaction (X×W):", round(slope_data$coefficients$b3, 3),
            "\nLow: b1 + b3×(-1 SD) =", round(slope_data$coefficients$b1 + slope_data$coefficients$b3 * -1, 3),
            "\nAvg: b1 =", round(slope_data$coefficients$b1, 3),
            "\nHigh: b1 + b3×(+1 SD) =", round(slope_data$coefficients$b1 + slope_data$coefficients$b3 * 1, 3)
          ),
          hjust = 1,
          vjust = -0.5,
          size = 3,
          color = "grey30"
        )
      
      # Save the plot
      ggsave(
        filename = file,
        plot = p,
        width = input$slope_plot_width / 100,
        height = input$slope_plot_height / 100,
        dpi = 300,
        bg = "white"
      )
      
      showNotification("Slope plot downloaded successfully!", type = "message")
    }
  )
  
  output$mediationInterpretation <- renderUI({
    req(model_fit())
    fit <- model_fit()
    
    if(is.null(fit)) {
      return(NULL)
    }
    
    tryCatch({
      if(input$use_bootstrap) {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
      } else {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
      }
      
      defined_effects <- pe[pe$op == ":=", ]
      
      if(nrow(defined_effects) == 0) {
        return(div(class = "interpretation-box",
                   h4("Results Interpretation"),
                   p("No defined effects found in the model. Please define effects using ':=' syntax in your model specification.")))
      }
      
      interpretations <- list()
      
      # Simple Mediation Interpretation (using user-friendly names)
      # Only show if using interactive interface OR if analysis_type is Simple Mediation
      if((current_interface() == "interactive" && input$analysis_type_interactive == "Simple Mediation") ||
         (current_interface() == "syntax" && input$analysis_type == "Simple Mediation")) {
        indirect_effect <- defined_effects[defined_effects$lhs == "Indirect_Effect_M", ]
        total_effect <- defined_effects[defined_effects$lhs == "Total_Effect", ]
        direct_effect <- pe[pe$op == "~" & pe$rhs == selected_variables$X & pe$lhs == selected_variables$Y, ]
        
        if(nrow(indirect_effect) > 0) {
          indirect_sig <- indirect_effect$pvalue[1] < 0.05
          direct_sig <- if(nrow(direct_effect) > 0) direct_effect$pvalue[1] < 0.05 else FALSE
          
          if(indirect_sig && !direct_sig) {
            mediation_type <- "Full Mediation"
            interpretation_text <- "The indirect effect is significant while the direct effect is not, indicating complete mediation."
          } else if(indirect_sig && direct_sig) {
            mediation_type <- "Partial Mediation"
            interpretation_text <- "Both indirect and direct effects are significant, indicating partial mediation."
          } else {
            mediation_type <- "No Mediation"
            interpretation_text <- "The indirect effect is not statistically significant, indicating no mediation effect."
          }
          
          interpretations[["Simple Mediation"]] <- paste0(
            "<strong>Mediation Type:</strong> ", mediation_type, "<br>",
            "<strong>Interpretation:</strong> ", interpretation_text, "<br>",
            "<strong>Indirect Effect (", selected_variables$X, " → ", selected_variables$M, " → ", selected_variables$Y, "):</strong> β = ", round(indirect_effect$std.all[1], 3), 
            ", p = ", format.pval(indirect_effect$pvalue[1], digits = 3), 
            if(indirect_sig) " (significant)" else " (not significant)", "<br>"
          )
          
          if(nrow(total_effect) > 0) {
            interpretations[["Simple Mediation"]] <- paste0(
              interpretations[["Simple Mediation"]],
              "<strong>Total Effect (", selected_variables$X, " → ", selected_variables$Y, "):</strong> β = ", round(total_effect$std.all[1], 3), 
              ", p = ", format.pval(total_effect$pvalue[1], digits = 3), "<br>"
            )
          }
        }
      }
      
      # Serial Mediation Interpretation - FIXED VERSION with user-friendly names
      # Only show if using interactive interface OR if analysis_type is Serial Mediation
      if((current_interface() == "interactive" && input$analysis_type_interactive == "Serial Mediation") ||
         (current_interface() == "syntax" && input$analysis_type == "Serial Mediation")) {
        indirect1 <- defined_effects[defined_effects$lhs == "Indirect_Effect_M1", ]
        indirect2 <- defined_effects[defined_effects$lhs == "Indirect_Effect_M2", ]
        indirect_serial <- defined_effects[defined_effects$lhs == "Indirect_Effect_Serial", ]
        total_indirect <- defined_effects[defined_effects$lhs == "Total_Indirect_Effect", ]
        total_effect <- defined_effects[defined_effects$lhs == "Total_Effect", ]
        
        if(nrow(indirect1) > 0 || nrow(indirect2) > 0 || nrow(indirect_serial) > 0) {
          interpretations[["Serial Mediation"]] <- "<strong>Serial Mediation Analysis:</strong><br>"
          
          sig_count <- 0
          if(nrow(indirect1) > 0 && indirect1$pvalue[1] < 0.05) sig_count <- sig_count + 1
          if(nrow(indirect2) > 0 && indirect2$pvalue[1] < 0.05) sig_count <- sig_count + 1
          if(nrow(indirect_serial) > 0 && indirect_serial$pvalue[1] < 0.05) sig_count <- sig_count + 1
          
          if(sig_count > 0) {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>Conclusion:</strong> Serial mediation is present with ", sig_count, 
              " significant indirect pathway(s).<br><br>"
            )
          } else {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>Conclusion:</strong> No significant serial mediation effects found.<br><br>"
            )
          }
          
          interpretations[["Serial Mediation"]] <- paste0(
            interpretations[["Serial Mediation"]],
            "<strong>Indirect Pathways:</strong><br>"
          )
          
          if(nrow(indirect1) > 0) {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>Indirect Effect (", selected_variables$X, " → ", selected_variables$M1, " → ", selected_variables$Y, "):</strong> β = ", round(indirect1$std.all[1], 3), 
              ", p = ", format.pval(indirect1$pvalue[1], digits = 3), 
              if(indirect1$pvalue[1] < 0.05) " (significant)" else " (not significant)", "<br>"
            )
          }
          
          if(nrow(indirect2) > 0) {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>Indirect Effect (", selected_variables$X, " → ", selected_variables$M2, " → ", selected_variables$Y, "):</strong> β = ", round(indirect2$std.all[1], 3), 
              ", p = ", format.pval(indirect2$pvalue[1], digits = 3), 
              if(indirect2$pvalue[1] < 0.05) " (significant)" else " (not significant)", "<br>"
            )
          }
          
          if(nrow(indirect_serial) > 0) {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>Serial Mediation (", selected_variables$X, " → ", selected_variables$M1, " → ", selected_variables$M2, " → ", selected_variables$Y, "):</strong> β = ", round(indirect_serial$std.all[1], 3), 
              ", p = ", format.pval(indirect_serial$pvalue[1], digits = 3), 
              if(indirect_serial$pvalue[1] < 0.05) " (significant)" else " (not significant)", "<br>"
            )
          }
          
          if(nrow(total_indirect) > 0) {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<br><strong>Total Indirect Effect:</strong> β = ", round(total_indirect$std.all[1], 3), 
              ", p = ", format.pval(total_indirect$pvalue[1], digits = 3), 
              if(total_indirect$pvalue[1] < 0.05) " (significant)" else " (not significant)", "<br>"
            )
          }
          
          if(nrow(total_effect) > 0) {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>Total Effect (", selected_variables$X, " → ", selected_variables$Y, "):</strong> β = ", round(total_effect$std.all[1], 3), 
              ", p = ", format.pval(total_effect$pvalue[1], digits = 3), "<br>"
            )
          }
        }
      }
      
      # Moderation Interpretation - ONLY for interactive interface
      # Check if using interactive interface and moderation is selected
      if(current_interface() == "interactive" && input$analysis_type_interactive == "Moderation") {
        
        # Create interaction term name for searching
        if(!is.null(selected_variables$X) && !is.null(selected_variables$W)) {
          x_var <- selected_variables$X
          w_var <- selected_variables$W
          x_clean <- gsub("[^a-zA-Z0-9_]", "", x_var)
          w_clean <- gsub("[^a-zA-Z0-9_]", "", w_var)
          interaction_name <- paste0(x_clean, "_", w_clean)
        } else {
          interaction_name <- "XW"
        }
        
        # Look for interaction effect
        interaction_effect <- pe[pe$op == "~" & grepl(interaction_name, pe$rhs), ]
        
        if (nrow(interaction_effect) > 0) {
          
          interaction_sig <- interaction_effect$pvalue[1] < 0.05
          
          # Get the simple slopes from the defined effects
          simple_slope_low <- defined_effects[defined_effects$lhs == "Simple_Slope_Low", ]
          simple_slope_avg <- defined_effects[defined_effects$lhs == "Simple_Slope_Avg", ]
          simple_slope_high <- defined_effects[defined_effects$lhs == "Simple_Slope_High", ]
          
          interpretations[["Moderation"]] <- paste0(
            "<strong>Moderation Analysis:</strong><br>",
            "<strong>Interaction Effect (", x_var, " × ", w_var, "):</strong> β = ", round(interaction_effect$std.all[1], 3), 
            ", p = ", format.pval(interaction_effect$pvalue[1], digits = 3), 
            if (interaction_sig) " (significant - moderation present)" else " (not significant - no moderation)", "<br>"
          )
          
          # Add simple slopes if interaction is significant
          if (interaction_sig) {
            interpretations[["Moderation"]] <- paste0(
              interpretations[["Moderation"]],
              "<br><strong>Simple Slopes Analysis (Effect of ", x_var, " on ", selected_variables$Y, " at different levels of ", w_var, "):</strong><br>"
            )
            
            # Low slope (-1 SD)
            if (nrow(simple_slope_low) > 0) {
              slope_sig <- simple_slope_low$pvalue[1] < 0.05
              interpretations[["Moderation"]] <- paste0(
                interpretations[["Moderation"]],
                "<strong>Low ", w_var, " (-1 SD):</strong> β = ", round(simple_slope_low$std.all[1], 3), 
                ", p = ", format.pval(simple_slope_low$pvalue[1], digits = 3), 
                if (slope_sig) " (significant)" else " (not significant)", "<br>"
              )
            }
            
            # Average slope (mean)
            if (nrow(simple_slope_avg) > 0) {
              slope_sig <- simple_slope_avg$pvalue[1] < 0.05
              interpretations[["Moderation"]] <- paste0(
                interpretations[["Moderation"]],
                "<strong>Average ", w_var, " (mean):</strong> β = ", round(simple_slope_avg$std.all[1], 3), 
                ", p = ", format.pval(simple_slope_avg$pvalue[1], digits = 3), 
                if (slope_sig) " (significant)" else " (not significant)", "<br>"
              )
            }
            
            # High slope (+1 SD)
            if (nrow(simple_slope_high) > 0) {
              slope_sig <- simple_slope_high$pvalue[1] < 0.05
              interpretations[["Moderation"]] <- paste0(
                interpretations[["Moderation"]],
                "<strong>High ", w_var, " (+1 SD):</strong> β = ", round(simple_slope_high$std.all[1], 3), 
                ", p = ", format.pval(simple_slope_high$pvalue[1], digits = 3), 
                if (slope_sig) " (significant)" else " (not significant)", "<br>"
              )
            }
          }
          
        } else {
          # Check if we can find the interaction in the model with a different pattern
          all_interactions <- pe[pe$op == "~" & grepl("_", pe$rhs), ]
          
          if (nrow(all_interactions) > 0) {
            # Found some interaction term
            interpretations[["Moderation"]] <- paste0(
              "<strong>Moderation Analysis:</strong><br>",
              "<strong>Note:</strong> Interaction term found in model: ", all_interactions$rhs[1], "<br>",
              "Check the path estimates table for moderation effects.<br>"
            )
          } else {
            interpretations[["Moderation"]] <- paste0(
              "<strong>Moderation Analysis:</strong><br>",
              "<strong>Note:</strong> No interaction term detected in the model.<br>",
              "If you intended to test moderation, make sure your model includes an interaction term (e.g., X_W).<br>",
              "The app automatically creates this term when you select variables for moderation analysis."
            )
          }
        }
      }
      
      # Combine all interpretations
      if(length(interpretations) > 0) {
        interpretation_text <- paste(unlist(interpretations), collapse = "<br>")
        
        div(class = "interpretation-box",
            h4("Results Interpretation"),
            p(HTML(interpretation_text))
        )
      } else {
        interpretation_text <- "<strong>Defined Effects Analysis:</strong><br>"
        for(i in 1:nrow(defined_effects)) {
          effect <- defined_effects[i, ]
          sig <- effect$pvalue < 0.05
          interpretation_text <- paste0(
            interpretation_text,
            "<strong>", effect$lhs, ":</strong> β = ", round(effect$std.all, 3), 
            ", p = ", format.pval(effect$pvalue, digits = 3), 
            if(sig) " (significant)" else " (not significant)", "<br>"
          )
        }
        
        div(class = "interpretation-box",
            h4("Results Interpretation"),
            p(HTML(interpretation_text))
        )
      }
      
    }, error = function(e) {
      div(class = "interpretation-box",
          h4("Results Interpretation"),
          p("Error generating interpretation. Please check your model specification.")
      )
    })
  })
  
  # Diagram title UI
  output$diagram_title_ui <- renderUI({
    if(nchar(input$diagram_title) > 0) {
      div(class = "diagram-title-container",
          h3(input$diagram_title, class = "diagram-title")
      )
    }
  })
  
  # Function to create enhanced diagram with coefficients and p-values
  createEnhancedDiagram <- function(fit, layout, node_size, text_size, edge_label_size, 
                                    arrow_size, edge_width, man_color, lat_color, 
                                    edge_color, node_labels) {
    
    if(input$use_bootstrap) {
      pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
    } else {
      pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
    }
    
    path_estimates <- pe[pe$op == "~", ]
    
    custom_edge_labels <- character()
    for(i in 1:nrow(path_estimates)) {
      est <- round(path_estimates$std.all[i], 3)
      pval <- path_estimates$pvalue[i]
      
      if (pval < 0.001) {
        p_label <- "p<0.001"
      } else if (pval < 0.01) {
        p_label <- sprintf("p=%.3f", pval)
      } else {
        p_label <- sprintf("p=%.3f", pval)
      }
      
      label <- sprintf("%.3f\n%s", est, p_label)
      custom_edge_labels <- c(custom_edge_labels, label)
    }
    
    p <- semPaths(fit, 
                  what = "std", 
                  layout = layout,
                  style = "lisrel",
                  residuals = FALSE, 
                  edge.label.cex = edge_label_size,
                  sizeMan = node_size,
                  sizeLat = node_size,
                  color = list(lat = lat_color, man = man_color),
                  edge.color = edge_color,
                  edge.width = edge_width,
                  node.width = 1.5, 
                  node.height = 1.5,
                  fade = FALSE,
                  edge.label.position = 0.6,
                  rotation = 2,
                  asize = arrow_size,
                  label.cex = text_size,
                  nodeLabels = node_labels,
                  edgeLabels = custom_edge_labels)
    
    return(p)
  }
  
  # Function to add mediation effects to PNG using magick
  addMediationEffectsToPNG <- function(png_file, fit) {
    if(input$use_bootstrap) {
      pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
    } else {
      pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
    }
    
    defined_effects <- pe[pe$op == ":=" & grepl("Indirect_Effect|Total_Effect|Simple_Slope", pe$lhs, ignore.case = TRUE), ]
    
    img <- image_read(png_file)
    
    info <- image_info(img)
    width <- info$width
    height <- info$height
    
    effects_text <- character()
    
    for(i in 1:nrow(defined_effects)) {
      effect_name <- defined_effects$lhs[i]
      std_coef <- round(defined_effects$std.all[i], 3)
      p_value <- defined_effects$pvalue[i]
      
      if (p_value < 0.001) {
        p_display <- "p<0.001"
      } else {
        p_display <- sprintf("p=%.3f", p_value)
      }
      
      if (effect_name == "Indirect_Effect_M") {
        display_name <- "Indirect Effect (X→M→Y)"
      } else if (effect_name == "Indirect_Effect_M1") {
        display_name <- "Indirect Effect (X→M1→Y)"
      } else if (effect_name == "Indirect_Effect_M2") {
        display_name <- "Indirect Effect (X→M2→Y)"
      } else if (effect_name == "Indirect_Effect_Serial") {
        display_name <- "Serial Mediation (X→M1→M2→Y)"
      } else if (effect_name == "Total_Indirect_Effect") {
        display_name <- "Total Indirect Effects"
      } else if (effect_name == "Total_Effect") {
        display_name <- "Total Effect"
      } else if (effect_name == "Simple_Slope_Low") {
        display_name <- "Simple Slope (Low)"
      } else if (effect_name == "Simple_Slope_Avg") {
        display_name <- "Simple Slope (Average)"
      } else if (effect_name == "Simple_Slope_High") {
        display_name <- "Simple Slope (High)"
      } else {
        display_name <- effect_name
      }
      
      effect_result <- sprintf("%s (β=%.3f, %s)", display_name, std_coef, p_display)
      effects_text <- c(effects_text, effect_result)
    }
    
    full_text <- paste(effects_text, collapse = "\n")
    title_text <- "Effect Estimates\n"
    full_text_with_title <- paste0(title_text, full_text)
    
    base_font_size <- min(width, height) / 45
    
    img_with_text <- image_annotate(img, full_text_with_title,
                                    location = "+20+20",
                                    size = base_font_size,
                                    color = "darkblue",
                                    weight = 700,
                                    boxcolor = "white",
                                    degrees = 0)
    
    image_write(img_with_text, png_file)
    
    return(png_file)
  }
  
  # Main diagram plot
  output$semPlot <- renderPlot({
    req(model_fit())
    fit <- model_fit()
    
    if(is.null(fit)) {
      return(NULL)
    }
    
    tryCatch({
      node_labels <- list()
      
      if(nchar(input$node_x_label) > 0) {
        node_labels[["X"]] <- substr(input$node_x_label, 1, 50)
      }
      if(nchar(input$node_y_label) > 0) {
        node_labels[["Y"]] <- substr(input$node_y_label, 1, 50)
      }
      if(nchar(input$node_m_label) > 0) {
        node_labels[["M"]] <- substr(input$node_m_label, 1, 50)
      }
      if(nchar(input$node_m1_label) > 0) {
        node_labels[["M1"]] <- substr(input$node_m1_label, 1, 50)
      }
      
      createEnhancedDiagram(fit, 
                            layout = input$diagram_layout,
                            node_size = input$node_size,
                            text_size = input$text_size,
                            edge_label_size = input$edge_label_size,
                            arrow_size = input$arrow_size,
                            edge_width = input$edge_width,
                            man_color = input$man_color,
                            lat_color = input$lat_color,
                            edge_color = input$edge_color,
                            node_labels = node_labels)
      
    }, error = function(e) {
      showNotification(paste("Error creating diagram:", e$message), type = "error")
      return(NULL)
    })
  }, height = function() {
    ifelse(is.null(input$main_tabs), 600, 
           ifelse(input$main_tabs == "Diagram", 
                  session$clientData$output_semPlot_width * 0.7, 400))
  })
  
  # Download diagram with enhanced features including magick text annotation
  output$download_diagram <- downloadHandler(
    filename = function() paste0("MedModr_Diagram_", Sys.Date(), ".png"),
    content = function(file) {
      req(model_fit())
      fit <- model_fit()
      
      if(is.null(fit)) {
        showNotification("Cannot download diagram - no valid model", type = "error")
        return()
      }
      
      temp_png <- tempfile(fileext = ".png")
      
      tryCatch({
        node_labels <- list()
        
        if(nchar(input$node_x_label) > 0) {
          node_labels[["X"]] <- substr(input$node_x_label, 1, 50)
        }
        if(nchar(input$node_y_label) > 0) {
          node_labels[["Y"]] <- substr(input$node_y_label, 1, 50)
        }
        if(nchar(input$node_m_label) > 0) {
          node_labels[["M"]] <- substr(input$node_m_label, 1, 50)
        }
        if(nchar(input$node_m1_label) > 0) {
          node_labels[["M1"]] <- substr(input$node_m1_label, 1, 50)
        }
        
        png(temp_png, width = input$diagram_width, height = input$diagram_height, res = 300)
        
        createEnhancedDiagram(fit, 
                              layout = input$diagram_layout,
                              node_size = input$node_size,
                              text_size = input$text_size,
                              edge_label_size = input$edge_label_size,
                              arrow_size = input$arrow_size,
                              edge_width = input$edge_width,
                              man_color = input$man_color,
                              lat_color = input$lat_color,
                              edge_color = input$edge_color,
                              node_labels = node_labels)
        
        if(nchar(input$diagram_title) > 0) {
          title(main = input$diagram_title, line = 1, cex.main = 2.5)
        }
        
        dev.off()
        
        final_png <- addMediationEffectsToPNG(temp_png, fit)
        
        file.copy(final_png, file, overwrite = TRUE)
        
      }, error = function(e) {
        showNotification(paste("Error creating downloadable diagram:", e$message), type = "error")
      })
      
      showNotification("Diagram downloaded successfully with coefficients, p-values, and mediation effects!", type = "message")
    }
  )
  
  # Keep the original download_plot for backward compatibility
  output$download_plot <- downloadHandler(
    filename = function() paste0("MedModr_Diagram_", Sys.Date(), ".png"),
    content = function(file) {
      output$download_diagram$func()(file)
    }
  )
  
  output$download_report <- downloadHandler(
    filename = function() paste0("MedModr_Report_", Sys.Date(), ".docx"),
    content = function(file) {
      tryCatch({
        fit <- model_fit()
        if(is.null(fit)) {
          showNotification("Cannot generate report - no valid model", type = "error")
          return(NULL)
        }
        
        # Determine analysis type based on interface
        if (current_interface() == "interactive") {
          analysis_type <- input$analysis_type_interactive
        } else {
          analysis_type <- input$analysis_type
        }
        
        # If analysis_type is NULL, try to detect from model
        if (is.null(analysis_type) || analysis_type == "") {
          # Try to detect from model syntax
          if (!is.null(input$model_syntax)) {
            if (grepl("Simple_Slope", input$model_syntax)) {
              analysis_type <- "Moderation"
            } else if (grepl("Indirect_Effect_M1", input$model_syntax) || grepl("Serial", input$model_syntax)) {
              analysis_type <- "Serial Mediation"
            } else if (grepl("Indirect_Effect_M", input$model_syntax)) {
              analysis_type <- "Simple Mediation"
            } else {
              analysis_type <- "Unknown"
            }
          } else {
            analysis_type <- "Unknown"
          }
        }
        
        # Get parameter estimates with proper error handling
        if(input$use_bootstrap) {
          pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
        } else {
          pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
        }
        
        # Create a temporary file for the diagram
        plot_file <- tempfile(pattern = "diagram_", fileext = ".png")
        
        # Create diagram with error handling
        tryCatch({
          png(plot_file, width = 1600, height = 1200, res = 300, type = "cairo-png")
          
          node_labels <- list()
          if(nchar(input$node_x_label) > 0) node_labels[["X"]] <- substr(input$node_x_label, 1, 50)
          if(nchar(input$node_y_label) > 0) node_labels[["Y"]] <- substr(input$node_y_label, 1, 50)
          if(nchar(input$node_m_label) > 0) node_labels[["M"]] <- substr(input$node_m_label, 1, 50)
          if(nchar(input$node_m1_label) > 0) node_labels[["M1"]] <- substr(input$node_m1_label, 1, 50)
          
          semPaths(fit, 
                   what = "std", 
                   layout = input$diagram_layout,
                   style = "lisrel",
                   residuals = FALSE, 
                   edge.label.cex = 1.2,
                   sizeMan = 10,
                   sizeLat = 10,
                   color = list(lat = input$lat_color, man = input$man_color),
                   edge.color = input$edge_color,
                   edge.width = 1.5,
                   node.width = 1.5, 
                   node.height = 1.5,
                   fade = FALSE,
                   edge.label.position = 0.6,
                   rotation = 2,
                   asize = 1,
                   label.cex = 1,
                   nodeLabels = node_labels)
          
          if(nchar(input$diagram_title) > 0) {
            title(main = input$diagram_title, line = 1, cex.main = 2.5)
          }
          dev.off()
        }, error = function(e) {
          if(file.exists(plot_file)) unlink(plot_file)
          png(plot_file, width = 800, height = 600, res = 100)
          plot(1, type="n", axes=FALSE, xlab="", ylab="")
          text(1, 1, "Diagram could not be generated\nCheck model specification", 
               cex = 1.2, col = "red")
          dev.off()
        })
        
        # Create the Word document
        doc <- read_docx() 
        
        # Add header
        doc <- doc %>%
          body_add_par(app_name, style = "heading 1") %>%
          body_add_par(paste("Analysis Report -", Sys.Date()), style = "heading 2")
        
        # Analysis Settings
        doc <- doc %>%
          body_add_par("Analysis Settings", style = "heading 3") %>%
          body_add_par(paste("Analysis type:", analysis_type), style = "Normal") %>%
          body_add_par(paste("Bootstrap samples:", ifelse(input$use_bootstrap, input$bootstrap_samples, "Not used")), 
                       style = "Normal") %>%
          body_add_par("Model Syntax", style = "heading 3") %>%
          body_add_par(input$model_syntax, style = "Normal")
        
        # Model Fit Indices
        doc <- doc %>%
          body_add_par("Model Fit Indices", style = "heading 3")
        
        fit_measures <- fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "rmsea.ci.lower", "rmsea.ci.upper", "srmr"))
        fit_text <- data.frame(
          Index = c("Chi-square", "df", "p-value", "CFI", "TLI", "RMSEA", "RMSEA CI Lower", "RMSEA CI Upper", "SRMR"),
          Value = c(
            sprintf("%.3f", fit_measures["chisq"]),
            sprintf("%.0f", fit_measures["df"]),
            format.pval(fit_measures["pvalue"], digits = 3),
            sprintf("%.3f", fit_measures["cfi"]),
            sprintf("%.3f", fit_measures["tli"]),
            sprintf("%.3f", fit_measures["rmsea"]),
            sprintf("%.3f", fit_measures["rmsea.ci.lower"]),
            sprintf("%.3f", fit_measures["rmsea.ci.upper"]),
            sprintf("%.3f", fit_measures["srmr"])
          )
        )
        
        ft <- flextable(fit_text) %>%
          theme_box() %>%
          autofit()
        doc <- body_add_flextable(doc, ft)
        
        # MODERATION ANALYSIS RESULTS SECTION (Using only available styles)
        if (analysis_type == "Moderation") {
          doc <- doc %>%
            body_add_par("Moderation Analysis Results", style = "heading 3")
          
          # Get all paths
          all_paths <- pe[pe$op == "~", ]
          
          # Identify variables
          if (current_interface() == "interactive" && !is.null(selected_variables$X) && !is.null(selected_variables$W)) {
            x_var <- selected_variables$X
            w_var <- selected_variables$W
            y_var <- selected_variables$Y
            x_clean <- gsub("[^a-zA-Z0-9_]", "", x_var)
            w_clean <- gsub("[^a-zA-Z0-9_]", "", w_var)
            interaction_name <- paste0(x_clean, "_", w_clean)
          } else {
            # Try to detect from model syntax
            x_var <- "X"
            w_var <- "W"
            y_var <- "Y"
            interaction_name <- "X_W"
          }
          
          # Find main effects and interaction
          main_effects <- all_paths[!grepl("_", all_paths$rhs) & !grepl(interaction_name, all_paths$rhs, ignore.case = TRUE), ]
          interaction_effect <- all_paths[grepl(interaction_name, all_paths$rhs, ignore.case = TRUE) | grepl("_", all_paths$rhs), ]
          
          # Get simple slopes from defined effects
          simple_slopes <- pe[pe$op == ":=" & grepl("Simple_Slope", pe$lhs, ignore.case = TRUE), ]
          
          # Table 1: Main Effects (using heading 3 for sub-sections)
          if (nrow(main_effects) > 0) {
            main_effects$Pathway <- paste(main_effects$lhs, "←", main_effects$rhs)
            main_table <- main_effects[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
            names(main_table) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
            main_table$`p-value` <- format.pval(main_table$`p-value`, digits = 3, eps = 0.001)
            main_table[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")] <- 
              round(main_table[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")], 3)
            
            doc <- doc %>%
              body_add_par("Main Effects", style = "heading 3") %>%
              body_add_par(paste("Effects of", x_var, "and", w_var, "on", y_var), style = "Normal")
            
            ft <- flextable(main_table) %>%
              theme_box() %>%
              autofit() %>%
              colformat_num(col_keys = c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper"), digits = 3) %>%
              bg(j = "p-value", 
                 bg = ifelse(as.numeric(main_table$`p-value`) < 0.001, "#FFE5E5",
                             ifelse(as.numeric(main_table$`p-value`) < 0.01, "#FFF0E5",
                                    ifelse(as.numeric(main_table$`p-value`) < 0.05, "#FFF5E5", "#F5F5F5"))))
            
            doc <- body_add_flextable(doc, ft)
          }
          
          # Table 2: Interaction Effect
          if (nrow(interaction_effect) > 0) {
            interaction_effect$Pathway <- paste(interaction_effect$lhs, "←", interaction_effect$rhs)
            int_table <- interaction_effect[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
            names(int_table) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
            int_table$`p-value` <- format.pval(int_table$`p-value`, digits = 3, eps = 0.001)
            int_table[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")] <- 
              round(int_table[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")], 3)
            
            doc <- doc %>%
              body_add_par("Interaction Effect", style = "heading 3") %>%
              body_add_par(paste("Interaction between", x_var, "and", w_var), style = "Normal")
            
            ft <- flextable(int_table) %>%
              theme_box() %>%
              autofit() %>%
              colformat_num(col_keys = c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper"), digits = 3) %>%
              bg(j = "p-value", 
                 bg = ifelse(as.numeric(int_table$`p-value`) < 0.001, "#FFE5E5",
                             ifelse(as.numeric(int_table$`p-value`) < 0.01, "#FFF0E5",
                                    ifelse(as.numeric(int_table$`p-value`) < 0.05, "#FFF5E5", "#F5F5F5"))))
            
            doc <- body_add_flextable(doc, ft)
            
            # Add interpretation of interaction
            interaction_sig <- as.numeric(interaction_effect$pvalue[1]) < 0.05
            interaction_text <- if(interaction_sig) {
              paste("✓ The interaction effect is statistically significant (p < .05), indicating that", 
                    w_var, "moderates the relationship between", x_var, "and", y_var, ".")
            } else {
              paste("✗ The interaction effect is not statistically significant (p > .05), indicating that", 
                    w_var, "does not moderate the relationship between", x_var, "and", y_var, ".")
            }
            
            doc <- doc %>%
              body_add_par(interaction_text, style = "Normal")
          }
          
          # Table 3: Simple Slopes Analysis
          if (nrow(simple_slopes) > 0) {
            slopes_table <- simple_slopes[, c("lhs", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
            # Clean up slope names
            slopes_table$lhs <- gsub("Simple_Slope_", "", slopes_table$lhs)
            slopes_table$lhs <- gsub("_", " ", slopes_table$lhs)
            slopes_table$lhs <- paste("Moderator:", slopes_table$lhs)
            
            names(slopes_table) <- c("Simple Slope", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
            slopes_table$`p-value` <- format.pval(slopes_table$`p-value`, digits = 3, eps = 0.001)
            slopes_table[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")] <- 
              round(slopes_table[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")], 3)
            
            doc <- doc %>%
              body_add_par("Simple Slopes Analysis", style = "heading 3") %>%
              body_add_par(paste("Effect of", x_var, "on", y_var, "at different levels of", w_var), style = "Normal")
            
            ft <- flextable(slopes_table) %>%
              theme_box() %>%
              autofit() %>%
              colformat_num(col_keys = c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper"), digits = 3) %>%
              bg(j = "p-value", 
                 bg = ifelse(as.numeric(slopes_table$`p-value`) < 0.001, "#FFE5E5",
                             ifelse(as.numeric(slopes_table$`p-value`) < 0.01, "#FFF0E5",
                                    ifelse(as.numeric(slopes_table$`p-value`) < 0.05, "#FFF5E5", "#F5F5F5"))))
            
            doc <- body_add_flextable(doc, ft)
            
            # Add simple slopes interpretation
            doc <- doc %>%
              body_add_par("Simple Slopes Interpretation:", style = "heading 3")
            
            for(i in 1:nrow(simple_slopes)) {
              slope_name <- gsub("Simple_Slope_", "", simple_slopes$lhs[i])
              slope_name <- gsub("_", " ", slope_name)
              slope_sig <- simple_slopes$pvalue[i] < 0.05
              slope_text <- if(slope_sig) {
                paste("• At", slope_name, "levels of", w_var, "the effect of", x_var, "on", y_var, 
                      "is statistically significant (β =", round(simple_slopes$std.all[i], 3), 
                      ", p =", format.pval(simple_slopes$pvalue[i], digits = 3), ").")
              } else {
                paste("• At", slope_name, "levels of", w_var, "the effect of", x_var, "on", y_var, 
                      "is not statistically significant (β =", round(simple_slopes$std.all[i], 3), 
                      ", p =", format.pval(simple_slopes$pvalue[i], digits = 3), ").")
              }
              doc <- body_add_par(doc, slope_text, style = "Normal")
            }
          }
          
          # Add moderator statistics if available
          if (!is.null(selected_variables$W) && !is.null(data())) {
            df <- data()
            if (selected_variables$W %in% names(df)) {
              w_data <- df[[selected_variables$W]]
              w_mean <- mean(w_data, na.rm = TRUE)
              w_sd <- sd(w_data, na.rm = TRUE)
              
              doc <- doc %>%
                body_add_par("Moderator Statistics", style = "heading 3") %>%
                body_add_par(paste("Mean of", selected_variables$W, ":", round(w_mean, 3)), style = "Normal") %>%
                body_add_par(paste("Standard Deviation of", selected_variables$W, ":", round(w_sd, 3)), style = "Normal") %>%
                body_add_par(paste("Low moderator level (-1 SD):", round(w_mean - w_sd, 3)), style = "Normal") %>%
                body_add_par(paste("High moderator level (+1 SD):", round(w_mean + w_sd, 3)), style = "Normal")
            }
          }
        }
        
        # For Mediation Analyses
        if (analysis_type %in% c("Simple Mediation", "Serial Mediation")) {
          doc <- doc %>%
            body_add_par(paste(analysis_type, "Results"), style = "heading 3")
          
          mediation_paths <- pe[pe$op == "~", ]
          indirect_effects <- pe[pe$op == ":=" & grepl("Indirect_Effect|Total_Effect", pe$lhs), ]
          relevant_paths <- rbind(mediation_paths, indirect_effects)
          
          relevant_paths$Pathway <- ifelse(relevant_paths$op == "~", 
                                           paste(relevant_paths$lhs, "←", relevant_paths$rhs),
                                           paste(relevant_paths$lhs, ":=", relevant_paths$rhs))
          
          table_data <- relevant_paths[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
          names(table_data) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
          
          table_data$`p-value` <- format.pval(table_data$`p-value`, digits = 3, eps = 0.001)
          table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")] <- 
            round(table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")], 3)
          
          ft <- flextable(table_data) %>%
            theme_box() %>%
            autofit() %>%
            colformat_num(col_keys = c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper"), digits = 3) %>%
            bg(j = "p-value", 
               bg = ifelse(as.numeric(table_data$`p-value`) < 0.001, "#FFE5E5",
                           ifelse(as.numeric(table_data$`p-value`) < 0.01, "#FFF0E5",
                                  ifelse(as.numeric(table_data$`p-value`) < 0.05, "#FFF5E5", "#F5F5F5"))))
          
          doc <- body_add_flextable(doc, ft)
        }
        
        # Add Path Diagram
        if(file.exists(plot_file) && file.info(plot_file)$size > 0) {
          doc <- doc %>%
            body_add_par("Path Diagram", style = "heading 3") %>%
            body_add_par("Note: Path coefficients show standardized estimates with p-values.", 
                         style = "Normal") %>%
            body_add_img(plot_file, width = 6, height = 4.5)
        }
        
        # Add Notes
        doc <- doc %>%
          body_add_par("Notes", style = "heading 3") %>%
          body_add_par("1. All estimates are based on maximum likelihood estimation using the lavaan package.", style = "Normal") %>%
          body_add_par("2. Standardized estimates (std.all) represent completely standardized solutions.", 
                       style = "Normal") %>%
          body_add_par("3. p-values < .05 are considered statistically significant.", 
                       style = "Normal") %>%
          body_add_par("4. For moderation analysis, simple slopes are calculated at -1 SD, mean, and +1 SD of the moderator.", 
                       style = "Normal") %>%
          body_add_par("Generated by MedModr", style = "Normal") %>%
          body_add_par(paste("Version:", app_version), style = "Normal") %>%
          body_add_par(paste("Release Date:", release_date), style = "Normal") %>%
          body_add_par(paste("Analysis Date:", format(Sys.Date(), "%B %d, %Y")), style = "Normal")
        
        # Save the document
        print(doc, target = file)
        
        # Clean up
        if(file.exists(plot_file)) unlink(plot_file)
        
        showNotification("Report generated successfully!", type = "message", duration = 5)
        
      }, error = function(e) {
        showNotification(paste("Error generating report:", e$message), 
                         type = "error", duration = 10)
        if(exists("plot_file") && file.exists(plot_file)) unlink(plot_file)
      })
    }
  )
  
}

shinyApp(ui, server)
