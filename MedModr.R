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
library(flextable)
library(shinythemes)
library(shinyjs)

app_name <- "MedModr"
app_version <- "1.2.0"

ui <- fluidPage(
  useShinyjs(),
  theme = shinytheme("cosmo"),
  tags$head(
    tags$style(HTML("
      @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500&family=Poppins:wght@600&display=swap');
      
      body {
        font-family: 'Roboto', sans-serif;
        background-color: #f5f7fa;
      }
      
      .app-title {
        font-family: 'Poppins', sans-serif;
        font-weight: 600;
        font-size: 32px;
        color: #2c3e50;
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 3px solid #3498db;
        background: linear-gradient(90deg, #3498db, #9b59b6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
      }
      
      .navbar {
        background-color: #2c3e50 !important;
        border-color: #2c3e50 !important;
      }
      
      .sidebar {
        background-color: #ffffff;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        padding: 20px;
        margin-right: 15px;
      }
      
      .main-panel {
        background-color: #ffffff;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        padding: 20px;
      }
      
      #fitText, #summaryText, #effectsText {
        font-family: 'Roboto', sans-serif;
        font-size: 14px;
        line-height: 1.5;
        white-space: pre-wrap;
        overflow-x: auto;
        resize: horizontal;
        min-height: 100px;
        width: 100%;
        border: 1px solid #e0e0e0;
        padding: 15px;
        background-color: #f8f9fa;
        border-radius: 6px;
      }
      
      .results-table {
        font-family: 'Roboto', sans-serif;
        font-size: 14px;
        width: 100%;
      }
      
      .well {
        background-color: #f8f9fa;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
      }
      
      .btn-primary {
        background-color: #3498db;
        border-color: #2980b9;
        color: white;
        font-weight: 500;
        border-radius: 6px;
        padding: 8px 16px;
        transition: all 0.3s ease;
      }
      
      .btn-primary:hover {
        background-color: #2980b9;
        border-color: #2472a4;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(41, 128, 185, 0.2);
      }
      
      .btn-success {
        background-color: #2ecc71;
        border-color: #27ae60;
        border-radius: 6px;
        transition: all 0.3s ease;
      }
      
      .btn-success:hover {
        background-color: #27ae60;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(46, 204, 113, 0.2);
      }
      
      .btn-info {
        background-color: #9b59b6;
        border-color: #8e44ad;
        border-radius: 6px;
        transition: all 0.3s ease;
      }
      
      .btn-info:hover {
        background-color: #8e44ad;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(155, 89, 182, 0.2);
      }
      
      .tab-content {
        padding: 20px;
        border-left: 1px solid #e0e0e0;
        border-right: 1px solid #e0e0e0;
        border-bottom: 1px solid #e0e0e0;
        border-radius: 0 0 8px 8px;
        background-color: #ffffff;
      }
      
      .nav-tabs > li > a {
        color: #7f8c8d;
        font-weight: 500;
        border-radius: 6px 6px 0 0;
      }
      
      .nav-tabs > li.active > a,
      .nav-tabs > li.active > a:hover,
      .nav-tabs > li.active > a:focus {
        color: #3498db;
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-bottom-color: transparent;
      }
      
      .diagram-controls {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 20px;
        border: 1px solid #e0e0e0;
      }
      
      .interpretation-box {
        background-color: #e8f4fc;
        border-left: 4px solid #3498db;
        padding: 15px;
        margin-top: 20px;
        margin-bottom: 20px;
        border-radius: 6px;
      }
      
      .instruction-box {
        background-color: #f0f8ff;
        border-left: 4px solid #9b59b6;
        padding: 15px;
        margin-bottom: 20px;
        border-radius: 6px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
      }
      
      h3, h4 {
        font-weight: 500;
        color: #2c3e50;
      }
      
      h3 {
        border-bottom: 2px solid #3498db;
        padding-bottom: 8px;
        margin-bottom: 15px;
      }
      
      h4 {
        color: #34495e;
      }
      
      .analysis-step {
        margin-bottom: 20px;
        padding: 15px;
        background-color: #f8f9fa;
        border-radius: 6px;
      }
      
      .analysis-step h5 {
        font-weight: 500;
        color: #2c3e50;
        margin-top: 0;
      }
      
      select, input, textarea {
        border-radius: 6px !important;
        border: 1px solid #e0e0e0 !important;
      }
      
      .file-input {
        padding: 10px;
        background-color: #f8f9fa;
        border-radius: 6px;
        border: 1px dashed #bdc3c7;
      }
      
      .shiny-input-container {
        margin-bottom: 15px;
      }
      
      .data-summary {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 6px;
        border: 1px solid #e0e0e0;
      }
    "))
  ),
  
  titlePanel(
    div(
      class = "app-title", app_name)
  ),
  
  sidebarLayout(
    sidebarPanel(
      width = 4,
      h4("Data Input", icon("database")),
      fileInput("datafile", "Upload your data (CSV, XLSX, Stata, or SPSS):", 
                accept = c(".csv", ".xlsx", ".xls", ".dta", ".sav", ".zsav", ".por")),
      
      uiOutput("var_preview"),
      
      h4("Analysis Settings", icon("cogs")),
      selectInput("analysis_type", "Select Analysis Type:",
                  choices = c("Simple Mediation", "Serial Mediation", "Moderation")),
      
      checkboxInput("use_composites", "Use Constructs (latent variables with indicators)?", value = FALSE),
      
      numericInput("bootstrap_samples", "Number of Bootstrap Samples:", 
                   value = 1000, min = 100, max = 10000, step = 100),
      
      checkboxInput("use_bootstrap", "Use Bootstrap Confidence Intervals?", value = FALSE),
      
      conditionalPanel(
        condition = "input.use_bootstrap == true",
        helpText("Note: Bootstrapping will take longer but provides more robust results.")
      ),
      
      # New UI element for covariates selection
      uiOutput("covariate_ui"),
      
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
        verbatimTextOutput("serial_mediation_syntax"),
        h4("Important Notes:"),
        tags$ul(
          tags$li("All mediators must be specified in order"),
          tags$li("Make sure your model syntax matches the serial mediation structure"),
          tags$li("Check that all variable names match your data")
        )
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
        verbatimTextOutput("moderation_syntax"),
        h4("Important Notes:"),
        tags$ul(
          tags$li("You must create the interaction term in your data first"),
          tags$li("Name it something like 'XW' or 'X_W'"),
          tags$li("Center your variables before creating interaction terms")
        )
      ),
      
      h4("Model Specification", icon("code")),
      textAreaInput("model_syntax", "Specify lavaan Model Syntax:", 
                    placeholder = "Enter your model syntax here...", height = "200px"),
      
      actionButton("run", "Run Analysis", icon = icon("play"), 
                   class = "btn-primary btn-lg"),
      br(), br(),
      downloadButton("download_report", "Download Word Report", class = "btn-success"),
      downloadButton("download_plot", "Download Diagram (PNG)", class = "btn-info")
    ),
    
    mainPanel(
      width = 8,
      tabsetPanel(
        tabPanel("Diagram", 
                 div(class = "diagram-controls",
                     h4("Diagram Formatting Options"),
                     fluidRow(
                       column(4,
                              selectInput("diagram_layout", "Layout:",
                                          choices = c("tree", "circle", "spring", "tree2", "tree3"),
                                          selected = "tree"),
                              sliderInput("node_size", "Node Size:", 
                                          min = 5, max = 20, value = 10, step = 1)
                       ),
                       column(4,
                              sliderInput("edge_label_size", "Edge Label Size:", 
                                          min = 0.5, max = 2, value = 1.2, step = 0.1),
                              sliderInput("arrow_size", "Arrow Size:", 
                                          min = 0.5, max = 2, value = 1, step = 0.1)
                       ),
                       column(4,
                              colourpicker::colourInput("man_color", "Observed Variable Color:", value = "lightyellow"),
                              colourpicker::colourInput("lat_color", "Latent Variable Color:", value = "skyblue")
                       )
                     )
                 ),
                 
                 plotOutput("semPlot", height = "600px"),
                 h4("Diagram Interpretation:"),
                 tags$ul(
                   tags$li("Arrows represent hypothesized relationships between variables"),
                   tags$li("Standardized coefficients are shown on the paths"),
                   tags$li("Solid lines indicate significant relationships (p < .05)"),
                   tags$li("Dashed lines indicate non-significant paths"),
                   tags$li("Rectangles represent observed variables"),
                   tags$li("Ovals represent latent variables (if used)")
                 )),
        tabPanel("Results", 
                 h3("Model Fit Indices", icon("check-circle")),
                 div(verbatimTextOutput("fitText"), style = "resize: both; overflow: auto;"),
                 
                 h3("Parameter Estimates Table", icon("table")),
                 DTOutput("estimatesTable"),
                 
                 h3("Effects Analysis", icon("project-diagram")),
                 div(verbatimTextOutput("effectsText"), style = "resize: both; overflow: auto;"),
                 
                 h3("Mediation Interpretation", icon("comment")),
                 uiOutput("mediationInterpretation"),
                 
                 h4("Results Interpretation:"),
                 tags$ul(
                   tags$li("Check model fit indices first (CFI > .90, RMSEA < .08 indicate good fit)"),
                   tags$li(paste("Bootstrap CIs are", 
                                 textOutput("bootstrap_status", inline = TRUE))),
                   tags$li("For mediation: Significant indirect effect (a*b) indicates mediation"),
                   tags$li("For serial mediation: Check all indirect paths (a1*a2*b2)"),
                   tags$li("For moderation: Significant interaction term (X:W) indicates moderation"),
                   tags$li("p-values < .05 are typically considered statistically significant"),
                   tags$li("Standardized coefficients (std.all) show effect sizes")
                 )),
        tabPanel("Variables in Data", 
                 DTOutput("var_table"),
                 br(),
                 h4("Data Summary"),
                 verbatimTextOutput("data_summary")),
        tabPanel("How to Run Analysis",
                 div(class = "well",
                     h3("Step-by-Step Analysis Guide", "v", app_version)),
                 
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
                         verbatimTextOutput("simple_mediation_indirect")),
                     div(class = "analysis-step",
                         h5("Step 4: Run analysis"),
                         p("Click 'Run Analysis' button to get results"))
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
                         verbatimTextOutput("serial_mediation_indirect")),
                     div(class = "analysis-step",
                         h5("Step 4: Run analysis"),
                         p("Click 'Run Analysis' button to get results"))
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
                         verbatimTextOutput("moderation_slopes")),
                     div(class = "analysis-step",
                         h5("Step 4: Run analysis"),
                         p("Click 'Run Analysis' button to get results"))
                 )
        ),
        tabPanel("About",
                 div(class = "well",
                     h4(app_name, "v", app_version),
                     p("A user-friendly interface for mediation, serial mediation, and moderation analysis using composite variables or constructs."),
                     h5("Key Features:"),
                     tags$ul(
                       tags$li("Supports simple and serial mediation models"),
                       tags$li("Moderation analysis with interaction terms"),
                       tags$li("Bootstrap confidence intervals"),
                       tags$li("Interactive path diagrams"),
                       tags$li("Professional report generation"),
                       tags$li("Adjust for confounding variables")
                     ),
                     h5("Developed by:"),
                     p("Mudasir Mohammed Ibrahim"),
                     h5("Contact:"),
                     p("mudassiribrahim30@gmail.com"),
                     h5("License:"),
                     p("MIT License - Free for academic and research use"),
                     h5("Citation:"),
                     p("Please cite this app if used in your research publications")
                 ))
      )
    )
  )
)

server <- function(input, output, session) {
  
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
  
  # Render UI for covariate selection
  output$covariate_ui <- renderUI({
    req(data())
    df <- data()
    var_names <- names(df)
    
    # Exclude potential key variables from covariates
    potential_keys <- c("X", "Y", "M", "M1", "M2", "W", "XW", "X_W", "interaction")
    covariate_choices <- setdiff(var_names, potential_keys)
    
    selectizeInput("covariates", "Select Covariates to Adjust For:",
                   choices = covariate_choices,
                   multiple = TRUE,
                   options = list(placeholder = 'Select variables to adjust for'))
  })
  
  # Variable preview
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
                "<br><strong>Observations:</strong> ", nrow(df)))
  })
  
  output$var_table <- renderDT({
    req(data())
    df <- data()
    if(ncol(df) > 500) {
      df <- df[, 1:500]  # Limit to first 500 variables
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
      df <- df[, 1:500]  # Limit to first 500 variables
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
  
  # Bootstrap status text
  output$bootstrap_status <- renderText({
    if(input$use_bootstrap) {
      paste("enabled with", input$bootstrap_samples, "samples")
    } else {
      "not used (using model-based standard errors)"
    }
  })
  
  # Syntax examples
  output$simple_mediation_syntax <- renderText({
    "M ~ a*X\nY ~ b*M + cp*X\n\n# Define effects\nindirect := a*b #Mediating effect\ntotal := cp + (a*b) #Total effect\n# Standardized indirect effect (Ignore, computed automatically)\nstd.indirect := std(a)*std(b)\n# Proportion mediated\nprop_mediated := (a*b)/total"
  })
  
  output$serial_mediation_syntax <- renderText({
    "Y ~ c*X\nM1 ~ a1*X\nM2 ~ a2*M1 + d*X\nY ~ b1*M1 + b2*M2\n\n# Define effects\nindirect1 := a1 * b1 #Simple mediation (for M1 only)\nindirect2 := a1 * a2 * b2 #Serial mediation (through M1 then M2)\nindirect3 := d * b2 #Simple mediation (for M2 only)\n\ntotal_indirect_effect := indirect1 + indirect2 + indirect3 #Total indirect effect\ntotal_effect := c + total_indirect_effect #Total effect"
  })
  
  output$moderation_syntax <- renderText({
    "# Note: You must create the interaction term in your data first\n# (e.g., data$XW <- data$X * data$W)\n\nY ~ b1*X + b2*W + b3*XW\n\n# Define simple slopes\nlowW := b1 + b3*(-1)\navgW := b1 + b3*(0)\nhighW := b1 + b3*(1)"
  })
  
  # How-to guide outputs
  output$simple_mediation_howto1 <- renderText({
    "M ~ a*X\nY ~ b*M + c*X"
  })
  
  output$simple_mediation_example <- renderText({
    "Anxiety ~ a*Stress_Composite\nAcademic_Performance ~ b*Anxiety + c*Stress_Composite"
  })
  
  output$simple_mediation_indirect <- renderText({
    "indirect := a*b"
  })
  
  output$serial_mediation_howto1 <- renderText({
    "Y ~ c*X\nM1 ~ a1*X\nM2 ~ a2*M1 + d*X\nY ~ b1*M1 + b2*M2"
  })
  
  output$serial_mediation_example <- renderText({
    "Life_Satisfaction ~ c*Career_calling\nAcademic_Motivation ~ a1*Career_calling\nStudy_Engagement ~ a2*Academic_Mot + d*Career_calling\nLife_Satisfaction ~ b1*Academic_Motivation + b2*Study_Engagement"
  })
  
  output$serial_mediation_indirect <- renderText({
    "indirect2 := a1 * a2 * b2"
  })
  
  output$moderation_howto1 <- renderText({
    "Y ~ b1*X + b2*W + b3*XW"
  })
  
  output$moderation_example <- renderText({
    "Academic_Performance ~ b1*Stress_Composite + b2*Coping_Construct + b3*Interaction"
  })
  
  output$moderation_slopes <- renderText({
    "lowW := b1 + b3*(-1)\navgW := b1 + b3*(0)\nhighW := b1 + b3*(1)"
  })
  
  # Model estimation with data size limits
  model_fit <- eventReactive(input$run, {
    req(input$model_syntax, data())
    df <- na.omit(data())
    
    # Apply size limits
    if(nrow(df) > 10000) {
      showNotification("Warning: Dataset has more than 10,000 observations. Only the first 10,000 will be used.", 
                       type = "warning", duration = 10)
      df <- df[1:10000, ]
    }
    
    if(ncol(df) > 500) {
      df <- df[, 1:500]  # Limit to first 500 variables
    }
    
    # For moderation analysis, check if interaction term exists
    if(input$analysis_type == "Moderation") {
      vars <- all.vars(parse(text = input$model_syntax))
      if(any(grepl("XW|X_W|interaction", vars, ignore.case = TRUE)) && 
         !any(grepl("XW|X_W|interaction", names(df), ignore.case = TRUE))) {
        showNotification("Error: Interaction term not found in data. Please create it first.", 
                         type = "error", duration = 10)
        return(NULL)
      }
    }
    
    # For serial mediation, validate model structure
    if(input$analysis_type == "Serial Mediation") {
      if(!grepl("M1.*X.*M2.*M1.*Y.*M2", input$model_syntax)) {
        showNotification("Warning: Serial mediation model should include paths from X to M1, M1 to M2, and M2 to Y", 
                         type = "warning", duration = 10)
      }
    }
    
    # Add covariates to model syntax if specified
    model_syntax <- input$model_syntax
    if(!is.null(input$covariates) && length(input$covariates) > 0) {
      # Add covariates to each regression equation
      cov_str <- paste(input$covariates, collapse = " + ")
      
      # Find all regression equations
      equations <- strsplit(model_syntax, "\n")[[1]]
      reg_equations <- grep("~", equations, value = TRUE)
      
      # Add covariates to each regression equation
      for(eq in reg_equations) {
        # Skip if already contains covariates
        if(!grepl(paste(input$covariates, collapse = "|"), eq)) {
          new_eq <- paste0(eq, " + ", cov_str)
          model_syntax <- gsub(eq, new_eq, model_syntax, fixed = TRUE)
        }
      }
    }
    
    tryCatch({
      if(input$use_bootstrap) {
        # Run with bootstrap
        sem(model = model_syntax, 
            data = df, 
            std.lv = TRUE,
            se = "bootstrap", 
            bootstrap = input$bootstrap_samples)
      } else {
        # Run without bootstrap
        sem(model = model_syntax, 
            data = df, 
            std.lv = TRUE)
      }
    }, error = function(e) {
      showNotification(paste("Error in model estimation:", e$message), 
                       type = "error", duration = 10)
      return(NULL)
    })
  })
  
  output$fitText <- renderPrint({
    req(model_fit())
    fit <- fitMeasures(model_fit(), c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "rmsea.ci.lower", "rmsea.ci.upper", "srmr"))
    cat("MODEL FIT INDICES:\n")
    cat(sprintf("Chi-square(%.0f) = %.3f, p = %.3f\n", fit["df"], fit["chisq"], fit["pvalue"]))
    cat(sprintf("CFI = %.3f, TLI = %.3f\n", fit["cfi"], fit["tli"]))
    cat(sprintf("RMSEA = %.3f, 90%% CI [%.3f, %.3f]\n", fit["rmsea"], fit["rmsea.ci.lower"], fit["rmsea.ci.upper"]))
    cat(sprintf("SRMR = %.3f\n\n", fit["srmr"]))
    
    cat("Fit Interpretation:\n")
    if (fit["cfi"] > 0.95 && fit["rmsea"] < 0.06) {
      cat("Excellent model fit (CFI > 0.95, RMSEA < 0.06)\n")
    } else if (fit["cfi"] > 0.90 && fit["rmsea"] < 0.08) {
      cat("Acceptable model fit (CFI > 0.90, RMSEA < 0.08)\n")
    } else {
      cat("Poor model fit - consider revising your model\n")
    }
    
    if(input$use_bootstrap) {
      cat(sprintf("\nBootstrap Results (based on %d samples):\n", input$bootstrap_samples))
    }
    
    # Show if covariates were included
    if(!is.null(input$covariates) && length(input$covariates) > 0) {
      cat("\nCovariates adjusted for in the analysis:\n")
      cat(paste(input$covariates, collapse = ", "), "\n")
    }
  })
  
  # Create formatted estimates table with SE
  output$estimatesTable <- renderDT({
    req(model_fit())
    
    if(input$use_bootstrap) {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE, 
                               boot.ci.type = "perc",
                               level = 0.95)
    } else {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE)
    }
    
    # Filter to show only relevant paths
    pe <- pe[pe$op %in% c("~", ":="), ]
    
    # Create pathway description
    pe$Pathway <- ifelse(pe$op == "~", 
                         paste(pe$lhs, "<-", pe$rhs),
                         paste(pe$lhs, ":=", pe$rhs))
    
    # Select and rename columns (now including SE)
    table_data <- pe[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
    names(table_data) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
    
    # Format p-values and numbers
    table_data$`p-value` <- format.pval(table_data$`p-value`, digits = 3, eps = 0.001)
    table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")] <- 
      round(table_data[, c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper")], 3)
    
    # Create datatable
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
  })
  
  output$effectsText <- renderPrint({
    req(model_fit())
    
    if(input$use_bootstrap) {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE, 
                               boot.ci.type = "perc",
                               level = 0.95)
    } else {
      pe <- parameterEstimates(model_fit(), 
                               standardized = TRUE, 
                               ci = TRUE)
    }
    
    # Defined effects (indirect, total, etc.)
    defined <- subset(pe, op == ":=")
    if (nrow(defined) > 0) {
      cat("DEFINED EFFECTS:\n\n")
      print(defined[, c("lhs", "op", "rhs", "est", "se", "pvalue", "ci.lower", "ci.upper")], 
            row.names = FALSE)
      
      # Add standardized indirect effects
      if (any(grepl("indirect", defined$lhs))) {
        std_effects <- defined[grepl("std.", defined$lhs, fixed = TRUE), ]
        if (nrow(std_effects) > 0) {
          cat("\nSTANDARDIZED EFFECTS:\n\n")
          print(std_effects[, c("lhs", "op", "rhs", "est", "se", "pvalue", "ci.lower", "ci.upper")], 
                row.names = FALSE)
        }
      }
    } else {
      cat("\nNo additional effects defined in the model.\n")
      cat("Tip: Define effects using ':=' syntax (e.g., 'indirect := a*b')\n")
    }
    
    if(input$use_bootstrap) {
      cat("\nNote: Effects estimates use bootstrap standard errors and confidence intervals\n")
    }
  })
  
  # Mediation interpretation
  output$mediationInterpretation <- renderUI({
    req(model_fit())
    
    if(input$analysis_type %in% c("Simple Mediation", "Serial Mediation")) {
      pe <- parameterEstimates(model_fit())
      
      # Get direct and indirect effects
      direct_effect <- pe[pe$lhs == "cp" | pe$rhs == "cp", "est"]
      indirect_effect <- pe[pe$lhs == "indirect", "est"]
      indirect_effect_p <- pe[pe$lhs == "indirect", "pvalue"]
      total_effect <- pe[pe$lhs == "total", "est"]
      
      if(length(direct_effect) == 0 || length(indirect_effect) == 0 || length(total_effect) == 0) {
        return(NULL)
      }
      
      # Calculate proportion mediated
      prop_mediated <- indirect_effect / total_effect
      
      # Determine mediation type
      mediation_type <- ifelse(
        abs(direct_effect) < 0.05 && indirect_effect_p < 0.05, 
        "Full Mediation",
        ifelse(
          indirect_effect_p < 0.05,
          "Partial Mediation",
          "No Mediation"
        )
      )
      
      # Create interpretation text
      interpretation <- if(mediation_type == "Full Mediation") {
        paste0(
          "The analysis suggests <strong>full mediation</strong> (indirect effect p = ", 
          format.pval(indirect_effect_p, digits = 3), 
          "). The direct effect is non-significant while the indirect effect is significant, ",
          "indicating that the mediator completely explains the relationship between X and Y."
        )
      } else if(mediation_type == "Partial Mediation") {
        paste0(
          "The analysis suggests <strong>partial mediation</strong> (indirect effect p = ", 
          format.pval(indirect_effect_p, digits = 3), 
          "). Both direct and indirect effects are significant, ",
          "indicating that the mediator partially explains the relationship between X and Y."
        )
      } else {
        paste0(
          "The analysis suggests <strong>no mediation</strong> (indirect effect p = ", 
          format.pval(indirect_effect_p, digits = 3), 
          "). The indirect effect is not statistically significant."
        )
      }
      
      # Create HTML output
      div(class = "interpretation-box",
          h4("Mediation Analysis Results:"),
          p(HTML(interpretation)),
          p(paste("Proportion mediated:", round(prop_mediated, 3)))
      )
    } else {
      return(NULL)
    }
  })
  
  output$semPlot <- renderPlot({
    req(model_fit())
    semPaths(model_fit(), 
             what = "std", 
             layout = input$diagram_layout,
             style = "lisrel",
             residuals = FALSE, 
             edge.label.cex = input$edge_label_size,
             sizeMan = input$node_size,
             sizeLat = input$node_size,
             color = list(lat = input$lat_color, man = input$man_color),
             edge.color = "black",
             node.width = 1.5, 
             node.height = 1.5,
             fade = FALSE,
             edge.label.position = 0.6,
             rotation = 2,
             asize = input$arrow_size)
  })
  
  output$download_report <- downloadHandler(
    filename = function() paste0("MedModr_Report_", Sys.Date(), ".docx"),
    content = function(file) {
      fit <- model_fit()
      if(is.null(fit)) {
        showNotification("Cannot generate report - no valid model", type = "error")
        return()
      }
      
      # Create a new Word document
      doc <- read_docx() 
      
      # Add title and basic info
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
      
      # Add model fit indices
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
      
      # Add parameter estimates (now with SE)
      if(input$use_bootstrap) {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE, boot.ci.type = "perc")
      } else {
        pe <- parameterEstimates(fit, standardized = TRUE, ci = TRUE)
      }
      
      pe <- pe[pe$op %in% c("~", ":="), ]
      pe$Pathway <- ifelse(pe$op == "~", 
                           paste(pe$lhs, "<-", pe$rhs),
                           paste(pe$lhs, ":=", pe$rhs))
      
      # Select and rename columns (including SE)
      table_data <- pe[, c("Pathway", "est", "std.all", "se", "pvalue", "ci.lower", "ci.upper")]
      names(table_data) <- c("Pathway", "Estimate", "Std. Estimate", "SE", "p-value", "CI Lower", "CI Upper")
      
      # Format p-values
      table_data$`p-value` <- format.pval(table_data$`p-value`, digits = 3, eps = 0.001)
      
      # Add parameter estimates table
      doc <- doc %>%
        body_add_par("Parameter Estimates", style = "heading 3")
      
      ft <- flextable(table_data) %>%
        theme_box() %>%
        autofit() %>%
        colformat_num(col_keys = c("Estimate", "Std. Estimate", "SE", "CI Lower", "CI Upper"), digits = 3) %>%
        bg(j = "p-value", 
           bg = ifelse(table_data$`p-value` < 0.001, "#FF6B6B",
                       ifelse(table_data$`p-value` < 0.01, "#FFA3A3",
                              ifelse(table_data$`p-value` < 0.05, "#FFD6A5", "#C8E7A5"))))
      
      doc <- body_add_flextable(doc, ft)
      
      # Add effects analysis
      defined <- subset(pe, op == ":=")
      if (nrow(defined) > 0) {
        doc <- doc %>%
          body_add_par("Defined Effects", style = "heading 3")
        
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
      
      # Add mediation interpretation if applicable
      if(input$analysis_type %in% c("Simple Mediation", "Serial Mediation")) {
        direct_effect <- pe[pe$lhs == "cp" | pe$rhs == "cp", "est"]
        indirect_effect <- pe[pe$lhs == "indirect", "est"]
        indirect_effect_p <- pe[pe$lhs == "indirect", "pvalue"]
        total_effect <- pe[pe$lhs == "total", "est"]
        
        if(length(direct_effect) > 0 && length(indirect_effect) > 0 && length(total_effect) > 0) {
          prop_mediated <- indirect_effect / total_effect
          
          mediation_type <- ifelse(
            abs(direct_effect) < 0.05 && indirect_effect_p < 0.05, 
            "Full Mediation",
            ifelse(
              indirect_effect_p < 0.05,
              "Partial Mediation",
              "No Mediation"
            )
          )
          
          interpretation <- if(mediation_type == "Full Mediation") {
            paste0(
              "The analysis suggests FULL MEDIATION (indirect effect p = ", 
              format.pval(indirect_effect_p, digits = 3), 
              "). The direct effect is non-significant while the indirect effect is significant, ",
              "indicating that the mediator completely explains the relationship between X and Y."
            )
          } else if(mediation_type == "Partial Mediation") {
            paste0(
              "The analysis suggests PARTIAL MEDIATION (indirect effect p = ", 
              format.pval(indirect_effect_p, digits = 3), 
              "). Both direct and indirect effects are significant, ",
              "indicating that the mediator partially explains the relationship between X and Y."
            )
          } else {
            paste0(
              "The analysis suggests NO MEDIATION (indirect effect p = ", 
              format.pval(indirect_effect_p, digits = 3), 
              "). The indirect effect is not statistically significant."
            )
          }
          
          doc <- doc %>%
            body_add_par("Mediation Interpretation", style = "heading 3") %>%
            body_add_par(interpretation, style = "Normal") %>%
            body_add_par(paste("Proportion mediated:", round(prop_mediated, 3)), style = "Normal")
        }
      }
      
      # Add covariates information if specified
      if(!is.null(input$covariates) && length(input$covariates) > 0) {
        doc <- doc %>%
          body_add_par("Covariates Adjusted For", style = "heading 3") %>%
          body_add_par(paste(input$covariates, collapse = ", "), style = "Normal")
      }
      
      # Add path diagram
      plot_file <- tempfile(fileext = ".png")
      png(plot_file, width = 800, height = 600)
      semPaths(fit, what = "std", layout = input$diagram_layout, style = "lisrel",
               residuals = FALSE, edge.label.cex = 1.0, 
               sizeMan = input$node_size, sizeLat = input$node_size,
               color = list(lat = input$lat_color, man = input$man_color),
               edge.color = "black", fade = FALSE, asize = input$arrow_size)
      dev.off()
      
      doc <- doc %>%
        body_add_par("Path Diagram", style = "heading 3") %>%
        body_add_img(plot_file, width = 6, height = 4.5) %>%
        
        body_add_par("Notes", style = "heading 3") %>%
        body_add_par("1. All estimates are based on maximum likelihood estimation.", style = "Normal") %>%
        body_add_par(paste("2. Bootstrap confidence intervals", 
                           ifelse(input$use_bootstrap, 
                                  paste("were used with", input$bootstrap_samples, "samples."),
                                  "were not used.")), 
                     style = "Normal") %>%
        body_add_par("3. Standardized estimates (std.all) represent completely standardized solutions.", 
                     style = "Normal") %>%
        body_add_par("4. p-values < .05 are considered statistically significant.", 
                     style = "Normal") %>%
        
        body_add_par("Generated by MedModr", style = "Normal") %>%
        body_add_par(paste("Version:", app_version), style = "Normal") %>%
        body_add_par(paste("Date:", format(Sys.Date(), "%B %d, %Y")), style = "Normal")
      
      # Save the document
      print(doc, target = file)
    }
  )
  
  output$download_plot <- downloadHandler(
    filename = function() paste0("MedModr_Diagram_", Sys.Date(), ".png"),
    content = function(file) {
      req(model_fit())
      png(file, width = 1000, height = 800)
      semPaths(model_fit(), 
               what = "std", 
               layout = input$diagram_layout,
               style = "lisrel",
               residuals = FALSE, 
               edge.label.cex = input$edge_label_size,
               sizeMan = input$node_size,
               sizeLat = input$node_size,
               color = list(lat = input$lat_color, man = input$man_color),
               edge.color = "black",
               fade = FALSE,
               asize = input$arrow_size)
      dev.off()
    }
  )
}

shinyApp(ui, server)
