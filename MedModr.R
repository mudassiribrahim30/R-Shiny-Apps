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

app_name <- "MedModr"
app_version <- "2.1.0"
release_date <- "October 2025"

ui <- fluidPage(
  useShinyjs(),
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
                     "Write lavaan syntax directly for maximum flexibility and control. For advanced users.")
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
          checkboxInput("use_composites_interactive", "Use Constructs (latent variables with indicators)?", value = FALSE),
          
          div(class = "analysis-options",
              h5("Output Options:"),
              checkboxGroupInput("output_options", "Select output to include:",
                                 choices = c("Total Effects" = "total",
                                             "Direct Effects" = "direct",
                                             "Indirect Effects" = "indirect",
                                             "Path Coefficients" = "paths",
                                             "Model Fit Indices" = "fit",
                                             "Bootstrap Results" = "bootstrap"),
                                 selected = c("total", "direct", "indirect", "paths", "fit"))
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
                      choices = c("Simple Mediation", "Serial Mediation", "Moderation")),
          
          checkboxInput("use_composites", "Use Constructs (latent variables with indicators)?", value = FALSE)
        ),
        
        # Common Analysis Settings
        checkboxInput("use_bootstrap", "Use Bootstrap Confidence Intervals?", value = FALSE),
        
        conditionalPanel(
          condition = "input.use_bootstrap == true",
          numericInput("bootstrap_samples", "Number of Bootstrap Samples:", 
                       value = 5000, min = 1000, max = 10000, step = 1000),
          div(class = "bootstrap-info",
              p("Note: By default, the app shows results without bootstrap. If you want bootstrap results, check this option and specify the number of samples.")
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
              ),
              
              conditionalPanel(
                condition = "input.analysis_type == 'Moderation'",
                h4("Moderation Guide:", icon("info-circle")),
                tags$ul(
                  tags$li("X = Independent Variable"),
                  tags$li("W = Moderator Variable"),
                  tags$li("Y = Dependent Variable"),
                  tags$li("XW = Interaction Term (X × W)"),
                  tags$li("Interpret significant interaction with simple slopes analysis")
                ),
                h4("Example Syntax:"),
                verbatimTextOutput("moderation_syntax")
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
                       h3("Model Fit Indices", icon("check-circle")),
                       div(verbatimTextOutput("fitText"), class = "resizable-text"),
                       
                       h3("Mediation Path Estimates (from Lavaan)", icon("table")),
                       DTOutput("mediationTable"),
                       
                       h3("Regression Results", icon("arrow-right")),
                       DTOutput("regressionTable"),
                       
                       h3("Effects Analysis", icon("project-diagram")),
                       div(verbatimTextOutput("effectsText"), class = "resizable-text"),
                       
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
                           h4("Diagram Controls", icon("sliders")),
                           actionButton("toggle_sidebar", "Show/Hide Data Panel", 
                                        icon = icon("bars"), 
                                        class = "btn-info btn-sm"),
                           br(), br(),
                           
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
                         column(
                           width = 9,
                           div(class = "diagram-main-area",
                               uiOutput("diagram_title_ui"),
                               plotOutput("semPlot", height = "100%")
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
                         tags$li("Improved bootstrap confidence intervals"),
                         tags$li("Interactive and customizable path diagrams"),
                         tags$li("Professional report generation"),
                         tags$li("Adjust for confounding variables")
                       ),
                       h5("Statistical Engine:"),
                       p("This app uses the", tags$strong("lavaan"), "package (Latent Variable Analysis) for all structural equation modeling analyses."),
                       h5("Developed by:"),
                       p("Mudasir Mohammed Ibrahim"),
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
  
  # Read data with support for larger files
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
        read_sav(input$datafile$datapath)
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
    selected_variables[[var_type]] <- input$variable_choice
    removeModal()
    
    runjs(paste0("$('#", tolower(var_type), "_box').addClass('has-variable');"))
  })
  
  # Generate model syntax from interactive selection - FIXED VERSION
  generate_model_syntax <- reactive({
    req(current_interface() == "interactive")
    
    if (input$analysis_type_interactive == "Simple Mediation") {
      req(selected_variables$X, selected_variables$M, selected_variables$Y)
      
      syntax <- paste0(
        "# Simple Mediation Model\n",
        selected_variables$M, " ~ a*", selected_variables$X, "\n",
        selected_variables$Y, " ~ b*", selected_variables$M, " + c*", selected_variables$X, "\n\n",
        "# Indirect and total effects\n",
        "indirect := a*b\n",
        "total := c + (a*b)\n",
        "prop_mediated := (a*b)/total"
      )
      
    } else if (input$analysis_type_interactive == "Serial Mediation") {
      req(selected_variables$X, selected_variables$M1, selected_variables$M2, selected_variables$Y)
      
      syntax <- paste0(
        "# Serial Mediation Model\n",
        selected_variables$M1, " ~ a1*", selected_variables$X, "\n",
        selected_variables$M2, " ~ a2*", selected_variables$M1, " + d*", selected_variables$X, "\n",
        selected_variables$Y, " ~ b1*", selected_variables$M1, " + b2*", selected_variables$M2, " + c*", selected_variables$X, "\n\n",
        "# Indirect effects - FIXED: Use valid lavaan labels\n",
        "indirect1 := a1 * b1\n",
        "indirect2 := a1 * a2 * b2\n", 
        "indirect3 := d * b2\n",
        "total_indirect := indirect1 + indirect2 + indirect3\n",
        "total_effect := c + total_indirect"
      )
      
    } else if (input$analysis_type_interactive == "Moderation") {
      req(selected_variables$X, selected_variables$W, selected_variables$Y)
      
      syntax <- paste0(
        "# Moderation Model\n",
        "# Note: You need to create the interaction term in your data first\n",
        "# Name it something like '", selected_variables$X, "_", selected_variables$W, "'\n",
        selected_variables$Y, " ~ b1*", selected_variables$X, " + b2*", selected_variables$W, " + b3*", selected_variables$X, "_", selected_variables$W, "\n\n",
        "# Simple slopes (assuming W is centered)\n",
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
      updateTextAreaInput(session, "model_syntax", value = generate_model_syntax())
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
                           tags$li("Note: You need to create the interaction term in your data first")
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
  
  output$var_table <- renderDT({
    req(data())
    df <- data()
    if(ncol(df) > 500) {
      df <- df[, 1:500]
    }
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
  
  # Render UI for covariate selection
  output$covariate_ui <- renderUI({
    req(data())
    df <- data()
    var_names <- names(df)
    
    potential_keys <- c("X", "Y", "M", "M1", "M2", "W", "XW", "X_W", "interaction")
    covariate_choices <- setdiff(var_names, potential_keys)
    
    selectizeInput("covariates", "Select Covariates to Adjust For:",
                   choices = covariate_choices,
                   multiple = TRUE,
                   options = list(placeholder = 'Select variables to adjust for'))
  })
  
  # Syntax examples - FIXED VERSION
  output$simple_mediation_syntax <- renderText({
    "# Replace X, M, Y with your actual variable names
M ~ a*X
Y ~ b*M + cp*X

# Define effects
indirect := a*b
total := cp + (a*b)
prop_mediated := (a*b)/total"
  })
  
  output$serial_mediation_syntax <- renderText({
    "# Replace X, M1, M2, Y with your actual variable names
M1 ~ a1*X
M2 ~ a2*M1 + d*X
Y ~ b1*M1 + b2*M2 + c*X

# Define effects - FIXED: Use valid lavaan labels
indirect1 := a1 * b1
indirect2 := a1 * a2 * b2
indirect3 := d * b2
total_indirect := indirect1 + indirect2 + indirect3
total_effect := c + total_indirect"
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
indirect := a*b
total := c + indirect"
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
    "# Add these effect definitions - FIXED: Use valid lavaan labels
indirect1 := a1 * b1
indirect2 := a1 * a2 * b2
indirect3 := d * b2
total_indirect := indirect1 + indirect2 + indirect3
total_effect := c + total_indirect"
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
                         "Serial_Mediation1", "Serial_Mediation2", "Serial_Mediation3", 
                         "Total_Indirect", "Total_Effect", "Simple_Slope_Low", 
                         "Simple_Slope_Avg", "Simple_Slope_High", "prop_mediated")
    
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
  
  # Enhanced model estimation with better error handling and syntax validation
  model_fit <- eventReactive(input$run, {
    req(input$model_syntax, data())
    
    # For interactive interface, check if variables are selected
    if (current_interface() == "interactive") {
      if (input$analysis_type_interactive == "Simple Mediation") {
        if (is.null(selected_variables$X) || is.null(selected_variables$M) || is.null(selected_variables$Y)) {
          showNotification("Please select all required variables for the analysis", type = "error")
          return(NULL)
        }
      } else if (input$analysis_type_interactive == "Serial Mediation") {
        if (is.null(selected_variables$X) || is.null(selected_variables$M1) || 
            is.null(selected_variables$M2) || is.null(selected_variables$Y)) {
          showNotification("Please select all required variables for the analysis", type = "error")
          return(NULL)
        }
      } else if (input$analysis_type_interactive == "Moderation") {
        if (is.null(selected_variables$X) || is.null(selected_variables$W) || is.null(selected_variables$Y)) {
          showNotification("Please select all required variables for the analysis", type = "error")
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
    
    df <- na.omit(df)
    
    if(nrow(df) < 10) {
      removeNotification("analysis_progress")
      showNotification("Error: Too few observations after removing missing values", type = "error", duration = 10)
      return(NULL)
    }
    
    if(nrow(df) > 10000) {
      showNotification("Warning: Dataset has more than 10,000 observations. Only the first 10,000 will be used.", 
                       type = "warning", duration = 10)
      df <- df[1:10000, ]
    }
    
    if(ncol(df) > 500) {
      df <- df[, 1:500]
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
    
    # Add covariates if specified
    if(!is.null(input$covariates) && length(input$covariates) > 0) {
      cov_str <- paste(input$covariates, collapse = " + ")
      
      equations <- strsplit(model_syntax, "\n")[[1]]
      reg_equations <- grep("~", equations, value = TRUE)
      
      for(i in seq_along(reg_equations)) {
        eq <- reg_equations[i]
        if(!grepl(paste(input$covariates, collapse = "|"), eq)) {
          new_eq <- paste0(eq, " + ", cov_str)
          model_syntax <- gsub(eq, new_eq, model_syntax, fixed = TRUE)
        }
      }
    }
    
    tryCatch({
      # Extract variables from syntax
      equations <- strsplit(model_syntax, "\n")[[1]]
      equations <- equations[!grepl("^\\s*#", equations)]
      equations <- equations[nchar(trimws(equations)) > 0]
      
      all_vars <- character()
      lavaan_keywords <- c("a", "b", "c", "cp", "d", "a1", "a2", "b1", "b2", "b3", 
                           "indirect", "total", "std", "lowW", "avgW", "highW",
                           "indirect1", "indirect2", "indirect3", "total_indirect", "total_effect",
                           "Simple_Slope_Low", "Simple_Slope_Avg", "Simple_Slope_High", "prop_mediated")
      
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
      
      data_vars <- unique(all_vars)
      available_vars <- names(df)
      missing_vars <- setdiff(data_vars, available_vars)
      
      if(length(missing_vars) > 0) {
        removeNotification("analysis_progress")
        showNotification(paste("Error: Variables not found in data:", paste(missing_vars, collapse = ", ")), 
                         type = "error", duration = 10)
        return(NULL)
      }
      
      fit_result <- NULL
      
      # Enhanced model estimation with better error handling
      if(input$use_bootstrap) {
        fit_result <- tryCatch({
          sem(model = model_syntax, 
              data = df, 
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
              data = df, 
              std.lv = TRUE,
              estimator = "ML",
              verbose = FALSE)
        })
      } else {
        fit_result <- sem(model = model_syntax, 
                          data = df, 
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
      if (fit_measures["cfi"] > 0.95 && fit_measures["rmsea"] < 0.06) {
        cat("Excellent model fit (CFI > 0.95, RMSEA < 0.06)\n")
      } else if (fit_measures["cfi"] > 0.90 && fit_measures["rmsea"] < 0.08) {
        cat("Acceptable model fit (CFI > 0.90, RMSEA < 0.08)\n")
      } else {
        cat("Poor model fit - consider revising your model\n")
      }
      
      if(input$use_bootstrap) {
        cat(sprintf("\nBootstrap Results (based on %d samples):\n", input$bootstrap_samples))
        cat("Note: Bootstrap provides more robust standard errors and confidence intervals\n")
      } else {
        cat("\nNote: Results are based on model-based standard errors (bootstrap not used)\n")
        cat("To use bootstrap, check 'Use Bootstrap Confidence Intervals' in Analysis Settings\n")
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
  
  # Create mediation path estimates table
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
      
      mediation_paths <- pe[pe$op == "~", ]
      indirect_effects <- pe[pe$op == ":=" & grepl("indirect|Serial_Mediation|Simple_Slope", pe$lhs), ]
      
      relevant_paths <- rbind(mediation_paths, indirect_effects)
      
      relevant_paths$Pathway <- ifelse(relevant_paths$op == "~", 
                                       paste(relevant_paths$lhs, "<-", relevant_paths$rhs),
                                       paste(relevant_paths$lhs, ":=", relevant_paths$rhs))
      
      table_data <- relevant_paths[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
      names(table_data) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
      
      table_data$`p-value` <- format.pval(table_data$`p-value`, digits = 3, eps = 0.001)
      table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")] <- 
        round(table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")], 3)
      
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
      
      defined <- subset(pe, op == ":=" & !grepl("indirect|Serial_Mediation|Simple_Slope", lhs))
      if (nrow(defined) > 0) {
        cat("ADDITIONAL DEFINED EFFECTS:\n\n")
        print(defined[, c("lhs", "op", "rhs", "est", "se", "pvalue", "ci.lower", "ci.upper")], 
              row.names = FALSE)
      } else {
        cat("\nNo additional effects defined in the model.\n")
        cat("Tip: Define effects using ':=' syntax (e.g., 'total := c + (a*b)')\n")
      }
      
      if(input$use_bootstrap) {
        cat("\nNote: Effects estimates use bootstrap standard errors and confidence intervals\n")
      } else {
        cat("\nNote: Effects estimates use model-based standard errors (bootstrap not used)\n")
      }
    }, error = function(e) {
      cat("Error generating effects text:", e$message, "\n")
    })
  })
  
  # Enhanced interpretation function for all analysis types - FIXED VERSION
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
      
      # Simple Mediation Interpretation
      if(input$analysis_type == "Simple Mediation" || input$analysis_type_interactive == "Simple Mediation") {
        indirect_effect <- defined_effects[defined_effects$lhs == "indirect", ]
        total_effect <- defined_effects[defined_effects$lhs == "total", ]
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
            "<strong>Indirect Effect:</strong> β = ", round(indirect_effect$std.all[1], 3), 
            ", p = ", format.pval(indirect_effect$pvalue[1], digits = 3), 
            if(indirect_sig) " (significant)" else " (not significant)", "<br>"
          )
          
          if(nrow(total_effect) > 0) {
            interpretations[["Simple Mediation"]] <- paste0(
              interpretations[["Simple Mediation"]],
              "<strong>Total Effect:</strong> β = ", round(total_effect$std.all[1], 3), 
              ", p = ", format.pval(total_effect$pvalue[1], digits = 3), "<br>"
            )
          }
        }
      }
      
      # Serial Mediation Interpretation - FIXED VERSION
      if(input$analysis_type == "Serial Mediation" || input$analysis_type_interactive == "Serial Mediation") {
        serial_effects <- defined_effects[grepl("indirect", defined_effects$lhs), ]
        total_indirect <- defined_effects[defined_effects$lhs %in% c("total_indirect", "total_indirect_effect"), ]
        total_effect <- defined_effects[defined_effects$lhs %in% c("total_effect", "total"), ]
        
        if(nrow(serial_effects) > 0) {
          interpretations[["Serial Mediation"]] <- "<strong>Serial Mediation Analysis:</strong><br>"
          
          sig_count <- sum(serial_effects$pvalue < 0.05)
          
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
          
          for(i in 1:nrow(serial_effects)) {
            effect <- serial_effects[i, ]
            sig <- effect$pvalue < 0.05
            
            if(effect$lhs == "indirect1") {
              path_desc <- "X → M1 → Y"
            } else if(effect$lhs == "indirect2") {
              path_desc <- "X → M1 → M2 → Y"
            } else if(effect$lhs == "indirect3") {
              path_desc <- "X → M2 → Y"
            } else {
              path_desc <- effect$lhs
            }
            
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>", path_desc, ":</strong> β = ", round(effect$std.all, 3), 
              ", p = ", format.pval(effect$pvalue, digits = 3), 
              if(sig) " (significant)" else " (not significant)", "<br>"
            )
          }
          
          if(nrow(total_indirect) > 0) {
            total_indirect_sig <- total_indirect$pvalue[1] < 0.05
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<br><strong>Total Indirect Effect:</strong> β = ", round(total_indirect$std.all[1], 3), 
              ", p = ", format.pval(total_indirect$pvalue[1], digits = 3), 
              if(total_indirect_sig) " (significant)" else " (not significant)", "<br>"
            )
          }
          
          if(nrow(total_effect) > 0) {
            interpretations[["Serial Mediation"]] <- paste0(
              interpretations[["Serial Mediation"]],
              "<strong>Total Effect:</strong> β = ", round(total_effect$std.all[1], 3), 
              ", p = ", format.pval(total_effect$pvalue[1], digits = 3), "<br>"
            )
          }
        }
      }
      
      # Moderation Interpretation
      if(input$analysis_type == "Moderation" || input$analysis_type_interactive == "Moderation") {
        interaction_effect <- pe[pe$op == "~" & grepl("XW", pe$rhs), ]
        simple_slopes <- defined_effects[grepl("Simple_Slope", defined_effects$lhs), ]
        
        if(nrow(interaction_effect) > 0) {
          interaction_sig <- interaction_effect$pvalue[1] < 0.05
          
          interpretations[["Moderation"]] <- paste0(
            "<strong>Moderation Analysis:</strong><br>",
            "<strong>Interaction Effect:</strong> β = ", round(interaction_effect$std.all[1], 3), 
            ", p = ", format.pval(interaction_effect$pvalue[1], digits = 3), 
            if(interaction_sig) " (significant - moderation present)" else " (not significant - no moderation)", "<br>"
          )
          
          if(interaction_sig && nrow(simple_slopes) > 0) {
            interpretations[["Moderation"]] <- paste0(
              interpretations[["Moderation"]],
              "<br><strong>Simple Slopes Analysis:</strong><br>"
            )
            
            for(i in 1:nrow(simple_slopes)) {
              slope <- simple_slopes[i, ]
              slope_sig <- slope$pvalue < 0.05
              
              if(slope$lhs == "Simple_Slope_Low") {
                level_desc <- "Low level of moderator"
              } else if(slope$lhs == "Simple_Slope_Avg") {
                level_desc <- "Average level of moderator"
              } else if(slope$lhs == "Simple_Slope_High") {
                level_desc <- "High level of moderator"
              } else {
                level_desc <- slope$lhs
              }
              
              interpretations[["Moderation"]] <- paste0(
                interpretations[["Moderation"]],
                "<strong>", level_desc, ":</strong> β = ", round(slope$std.all, 3), 
                ", p = ", format.pval(slope$pvalue, digits = 3), 
                if(slope_sig) " (significant)" else " (not significant)", "<br>"
              )
            }
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
    
    defined_effects <- pe[pe$op == ":=" & grepl("indirect|Serial_Mediation|Simple_Slope|Total", pe$lhs, ignore.case = TRUE), ]
    
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
      
      if (effect_name == "indirect") {
        display_name <- "Simple Mediation (X→M→Y)"
      } else if (effect_name == "indirect1") {
        display_name <- "Serial Mediation 1 (X→M1→Y)"
      } else if (effect_name == "indirect2") {
        display_name <- "Serial Mediation 2 (X→M1→M2→Y)"
      } else if (effect_name == "indirect3") {
        display_name <- "Serial Mediation 3 (X→M2→Y)"
      } else if (effect_name == "total_indirect") {
        display_name <- "Total Indirect Effects"
      } else if (effect_name == "total_effect") {
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
      fit <- model_fit()
      if(is.null(fit)) {
        showNotification("Cannot generate report - no valid model", type = "error")
        return()
      }
      
      doc <- read_docx() 
      
      doc <- doc %>%
        body_add_par(app_name, style = "heading 1") %>%
        body_add_par(paste("Analysis Report -", Sys.Date()), style = "heading 2") %>%
        body_add_par("Analysis Settings", style = "heading 3") %>%
        body_add_par(paste("Analysis type:", input$analysis_type), style = "Normal") %>%
        body_add_par(paste("Bootstrap samples:", ifelse(input$use_bootstrap, input$bootstrap_samples, "Not used")), 
                     style = "Normal") %>%
        body_add_par("Model Syntax", style = "heading 3") %>%
        body_add_par(input$model_syntax, style = "Normal") %>%
        body_add_par("Model Fit Indices", style = "heading 3")
      
      fit_measures <- fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr"))
      fit_text <- data.frame(
        Index = c("Chi-square", "df", "p-value", "CFI", "TLI", "RMSEA", "SRMR"),
        Value = c(
          sprintf("%.3f", fit_measures["chisq"]),
          sprintf("%.0f", fit_measures["df"]),
          format.pval(fit_measures["pvalue"], digits = 3),
          sprintf("%.3f", fit_measures["cfi"]),
          sprintf("%.3f", fit_measures["tli"]),
          sprintf("%.3f", fit_measures["rmsea"]),
          sprintf("%.3f", fit_measures["srmr"])
        )
      )
      
      ft <- flextable(fit_text) %>%
        theme_box() %>%
        autofit()
      doc <- body_add_flextable(doc, ft)
      
      if(input$use_bootstrap) {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
      } else {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
      }
      
      mediation_paths <- pe[pe$op == "~", ]
      indirect_effects <- pe[pe$op == ":=" & grepl("indirect|Serial_Mediation|Simple_Slope", pe$lhs), ]
      relevant_paths <- rbind(mediation_paths, indirect_effects)
      
      relevant_paths$Pathway <- ifelse(relevant_paths$op == "~", 
                                       paste(relevant_paths$lhs, "<-", relevant_paths$rhs),
                                       paste(relevant_paths$lhs, ":=", relevant_paths$rhs))
      
      table_data <- relevant_paths[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
      names(table_data) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
      
      table_data$`p-value` <- format.pval(table_data$`p-value`, digits = 3, eps = 0.001)
      
      doc <- doc %>%
        body_add_par("Mediation Path Estimates (from Lavaan)", style = "heading 3")
      
      ft <- flextable(table_data) %>%
        theme_box() %>%
        autofit() %>%
        colformat_num(col_keys = c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper"), digits = 3) %>%
        bg(j = "p-value", 
           bg = ifelse(table_data$`p-value` < 0.001, "#FF6B6B",
                       ifelse(table_data$`p-value` < 0.01, "#FFA3A3",
                              ifelse(table_data$`p-value` < 0.05, "#FFD6A5", "#C8E7A5"))))
      
      doc <- body_add_flextable(doc, ft)
      
      reg_results <- regression_results()
      if(!is.null(reg_results) && length(reg_results) > 0) {
        doc <- doc %>%
          body_add_par("Regression Results", style = "heading 3")
        
        reg_table_data <- data.frame()
        
        for(key in names(reg_results)) {
          res <- reg_results[[key]]
          
          unstd_row <- data.frame(
            Regression = res$formula,
            `Coefficient Type` = "Unstandardized",
            Estimate = sprintf("%.3f", res$unstandardized$estimate),
            SE = sprintf("%.3f", res$unstandardized$se),
            `CI Lower` = sprintf("%.3f", res$unstandardized$ci_lower),
            `CI Upper` = sprintf("%.3f", res$unstandardized$ci_upper),
            `p-value` = sprintf("%.3f", res$unstandardized$p_value),
            `R²` = sprintf("%.3f", res$r_squared),
            N = res$n_obs,
            check.names = FALSE
          )
          
          std_row <- data.frame(
            Regression = res$formula,
            `Coefficient Type` = "Standardized",
            Estimate = sprintf("%.3f", res$standardized$estimate),
            SE = sprintf("%.3f", res$standardized$se),
            `CI Lower` = sprintf("%.3f", res$standardized$ci_lower),
            `CI Upper` = sprintf("%.3f", res$standardized$ci_upper),
            `p-value` = sprintf("%.3f", res$standardized$p_value),
            `R²` = sprintf("%.3f", res$r_squared),
            N = res$n_obs,
            check.names = FALSE
          )
          
          reg_table_data <- rbind(reg_table_data, unstd_row, std_row)
        }
        
        ft <- flextable(reg_table_data) %>%
          theme_box() %>%
          autofit() %>%
          bg(j = "p-value", 
             bg = ifelse(as.numeric(reg_table_data$`p-value`) < 0.001, "#FF6B6B",
                         ifelse(as.numeric(reg_table_data$`p-value`) < 0.01, "#FFA3A3",
                                ifelse(as.numeric(reg_table_data$`p-value`) < 0.05, "#FFD6A5", "#C8E7A5"))))
        
        doc <- body_add_flextable(doc, ft)
      }
      
      defined <- subset(pe, op == ":=" & !grepl("indirect|Serial_Mediation|Simple_Slope", lhs))
      if (nrow(defined) > 0) {
        doc <- doc %>%
          body_add_par("Additional Defined Effects", style = "heading 3")
        
        effects_data <- defined[, c("lhs", "est", "se", "pvalue", "ci.lower", "ci.upper")]
        names(effects_data) <- c("Effect", "Estimate", "SE", "p-value", "CI Lower", "CI Upper")
        effects_data$`p-value` <- format.pval(effects_data$`p-value`, digits = 3)
        
        ft <- flextable(effects_data) %>%
          theme_box() %>%
          autofit() %>%
          colformat_num(col_keys = c("Estimate", "SE", "CI Lower", "CI Upper"), digits = 3) %>%
          bg(j = "p-value", 
             bg = ifelse(effects_data$`p-value` < 0.001, "#FF6B6B",
                         ifelse(effects_data$`p-value` < 0.01, "#FFA3A3",
                                ifelse(effects_data$`p-value` < 0.05, "#FFD6A5", "#C8E7A5"))))
        
        doc <- body_add_flextable(doc, ft)
      }
      
      if(input$analysis_type == "Simple Mediation" || input$analysis_type_interactive == "Simple Mediation") {
        indirect_effect <- pe[pe$lhs == "indirect", ]
        total_effect <- pe[pe$lhs == "total", ]
        
        if(nrow(indirect_effect) > 0) {
          interpretation_text <- paste0(
            "Simple Mediation Analysis:\n",
            "Indirect Effect: β = ", round(indirect_effect$std.all[1], 3), 
            ", p = ", format.pval(indirect_effect$pvalue[1], digits = 3),
            if(indirect_effect$pvalue[1] < 0.05) " (significant)" else " (not significant)", "\n",
            if(nrow(total_effect) > 0) paste0("Total Effect: β = ", round(total_effect$std.all[1], 3)) else ""
          )
          
          doc <- doc %>%
            body_add_par("Results Interpretation", style = "heading 3") %>%
            body_add_par(interpretation_text, style = "Normal")
        }
      } else if(input$analysis_type == "Serial Mediation" || input$analysis_type_interactive == "Serial Mediation") {
        serial_effects <- pe[grepl("indirect", pe$lhs) & pe$op == ":=", ]
        
        if(nrow(serial_effects) > 0) {
          interpretation_text <- "Serial Mediation Analysis:\n"
          for(i in 1:nrow(serial_effects)) {
            effect <- serial_effects[i, ]
            interpretation_text <- paste0(
              interpretation_text,
              effect$lhs, ": β = ", round(effect$std.all, 3), 
              ", p = ", format.pval(effect$pvalue, digits = 3),
              if(effect$pvalue < 0.05) " (significant)" else " (not significant)", "\n"
            )
          }
          
          doc <- doc %>%
            body_add_par("Results Interpretation", style = "heading 3") %>%
            body_add_par(interpretation_text, style = "Normal")
        }
      } else if(input$analysis_type == "Moderation" || input$analysis_type_interactive == "Moderation") {
        interaction_effect <- pe[pe$op == "~" & grepl("XW", pe$rhs), ]
        
        if(nrow(interaction_effect) > 0) {
          interpretation_text <- paste0(
            "Moderation Analysis:\n",
            "Interaction Effect: β = ", round(interaction_effect$std.all[1], 3), 
            ", p = ", format.pval(interaction_effect$pvalue[1], digits = 3),
            if(interaction_effect$pvalue[1] < 0.05) " (significant - moderation present)" else " (not significant - no moderation)"
          )
          
          doc <- doc %>%
            body_add_par("Results Interpretation", style = "heading 3") %>%
            body_add_par(interpretation_text, style = "Normal")
        }
      }
      
      if(!is.null(input$covariates) && length(input$covariates) > 0) {
        doc <- doc %>%
          body_add_par("Covariates Adjusted For", style = "heading 3") %>%
          body_add_par(paste(input$covariates, collapse = ", "), style = "Normal")
      }
      
      plot_file <- tempfile(fileext = ".png")
      png(plot_file, width = 1600, height = 1200, res = 300)
      
      createEnhancedDiagram(fit, 
                            layout = input$diagram_layout,
                            node_size = 10,
                            text_size = 1,
                            edge_label_size = 1.2,
                            arrow_size = 1,
                            edge_width = 1.5,
                            man_color = input$man_color,
                            lat_color = input$lat_color,
                            edge_color = input$edge_color,
                            node_labels = list())
      
      if(nchar(input$diagram_title) > 0) {
        title(main = input$diagram_title, line = 1, cex.main = 2.5)
      }
      dev.off()
      
      plot_file_with_effects <- addMediationEffectsToPNG(plot_file, fit)
      
      doc <- doc %>%
        body_add_par("Path Diagram", style = "heading 3") %>%
        body_add_par("Note: Path coefficients show standardized estimates with p-values. Effect estimates are shown in the bottom-left corner.", 
                     style = "Normal") %>%
        body_add_img(plot_file_with_effects, width = 6, height = 4.5) %>%
        
        body_add_par("Notes", style = "heading 3") %>%
        body_add_par("1. All estimates are based on maximum likelihood estimation using the lavaan package.", style = "Normal") %>%
        body_add_par(paste("2. Bootstrap confidence intervals", 
                           ifelse(input$use_bootstrap, 
                                  paste("were used with", input$bootstrap_samples, "samples."),
                                  "were not used.")), 
                     style = "Normal") %>%
        body_add_par("3. Standardized estimates (std.all) represent completely standardized solutions.", 
                     style = "Normal") %>%
        body_add_par("4. Regression results show all possible bivariate relationships between variables in the model.", 
                     style = "Normal") %>%
        body_add_par("5. p-values < .05 are considered statistically significant.", 
                     style = "Normal") %>%
        body_add_par("6. Path diagram shows standardized coefficients with p-values on separate lines.", 
                     style = "Normal") %>%
        
        body_add_par("Generated by MedModr", style = "Normal") %>%
        body_add_par(paste("Version:", app_version), style = "Normal") %>%
        body_add_par(paste("Release Date:", release_date), style = "Normal") %>%
        body_add_par(paste("Analysis Date:", format(Sys.Date(), "%B %d, %Y")), style = "Normal")
      
      print(doc, target = file)
    }
  )
}

shinyApp(ui, server)
