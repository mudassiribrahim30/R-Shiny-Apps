library(shiny)
library(shinydashboard)
library(haven)
library(readxl)
library(foreign)
library(psych)
library(GPArotation)
library(ggplot2)
library(shinythemes)
library(DT)
library(readr)
library(plotly)
library(dplyr)
library(shinyjs)

ui <- dashboardPage(
  dashboardHeader(
    title = "FactorGuard",
    titleWidth = 250
  ),
  dashboardSidebar(
    width = 300,
    collapsed = TRUE,
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
        
        /* Sidebar styling */
        .sidebar {
          background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%) !important;
          border-right: 1px solid var(--border-gray);
          box-shadow: 2px 0 10px rgba(0,0,0,0.1);
        }
        
        .sidebar-menu {
          background: transparent !important;
        }
        
        .sidebar-menu li a {
          font-family: 'Roboto', sans-serif;
          font-weight: 500;
          color: var(--text-dark) !important;
          border-left: 4px solid transparent;
          transition: all 0.3s ease;
          margin: 5px 10px;
          border-radius: 8px;
        }
        
        .sidebar-menu li a:hover {
          background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%) !important;
          color: var(--primary-blue) !important;
          border-left: 4px solid var(--primary-blue);
        }
        
        .sidebar-menu li.active a {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%) !important;
          color: white !important;
          border-left: 4px solid var(--primary-blue);
          box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
        }
        
        /* Main content styling */
        .content-wrapper, .right-side {
          background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%) !important;
          background-attachment: fixed;
          height: 100vh;
          overflow: hidden;
        }
        
        /* Main content area - fixed height with scrollable results */
        .main-content-container {
          height: calc(100vh - 80px);
          display: flex;
          flex-direction: column;
        }
        
        .fixed-sidebar {
          height: calc(100vh - 80px);
          overflow-y: auto;
        }
        
        .scrollable-results {
          height: calc(100vh - 80px);
          overflow-y: auto;
          background: linear-gradient(180deg, #ffffff 0%, #fafbfc 100%);
          border-radius: 15px;
          box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
          padding: 25px;
          border: 1px solid rgba(255,255,255,0.2);
        }
        
        .box {
          background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
          border-radius: 15px;
          border: none;
          box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
          margin-bottom: 20px;
          border: 1px solid rgba(255,255,255,0.2);
        }
        
        /* Tab styling */
        .nav-tabs-custom {
          background: transparent;
          border: none;
          box-shadow: none;
        }
        
        .nav-tabs-custom > .nav-tabs {
          background: transparent;
          border-bottom: 2px solid var(--border-gray);
          margin-bottom: 20px;
        }
        
        .nav-tabs-custom > .nav-tabs > li.active > a {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
          color: white;
          border: none;
          border-radius: 10px 10px 0 0;
          box-shadow: 0 -4px 12px rgba(52, 152, 219, 0.3);
          font-weight: 600;
        }
        
        .nav-tabs-custom > .nav-tabs > li > a {
          color: var(--text-medium);
          border: none;
          border-radius: 10px 10px 0 0;
          margin-right: 5px;
          transition: all 0.3s ease;
          font-weight: 500;
        }
        
        .nav-tabs-custom > .nav-tabs > li > a:hover {
          background: linear-gradient(135deg, var(--light-gray) 0%, #e9ecef 100%);
          color: var(--primary-blue);
        }
        
        .tab-content {
          background: transparent;
          border-radius: 0 15px 15px 15px;
          padding: 0;
          border: none;
        }
        
        /* Button styling */
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
        
        /* Input styling */
        .form-control, .selectize-input {
          border-radius: 8px !important;
          border: 1px solid var(--border-gray) !important;
          transition: all 0.3s ease;
          padding: 10px 12px;
        }
        
        .form-control:focus, .selectize-input.focus {
          border-color: var(--primary-blue) !important;
          box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1) !important;
          outline: none;
        }
        
        /* Well panel styling */
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
        
        /* File input styling */
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
        
        /* Syntax help boxes */
        .syntax-help {
          background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%);
          border-left: 4px solid var(--primary-blue);
          padding: 15px;
          margin: 15px 0;
          border-radius: 8px;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        }
        
        /* Instruction boxes */
        .instruction-box {
          background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%);
          border-left: 4px solid var(--primary-blue);
          padding: 20px;
          margin-bottom: 25px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        }
        
        /* Headers */
        h1, h2, h3, h4 {
          font-family: 'Poppins', sans-serif;
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
        
        /* DataTables styling */
        .dataTables_wrapper {
          background: white;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.05);
          padding: 15px;
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
        
        /* Footer styling */
        .footer {
          text-align: center;
          padding: 20px;
          margin-top: 30px;
          color: var(--text-medium);
          font-size: 14px;
          border-top: 1px solid var(--border-gray);
        }
        
        /* Loading animation */
        .shiny-notification {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
          color: white;
          border-radius: 10px;
          box-shadow: 0 8px 25px rgba(0,0,0,0.15);
          border: none;
        }
        
        /* Results content styling */
        .results-content {
          max-height: 100%;
        }
        
        .results-section {
          margin-bottom: 30px;
          padding: 20px;
          background: white;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.05);
          border: 1px solid var(--border-gray);
        }
        
        /* Download button container */
        .download-container {
          margin-top: 20px;
          padding: 15px;
          background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
          border-radius: 10px;
          border: 1px solid var(--border-gray);
          text-align: center;
        }
        
        /* Plot container */
        .plot-container {
          margin-bottom: 20px;
          border: 1px solid var(--border-gray);
          border-radius: 10px;
          overflow: hidden;
          box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        
        /* Developer info styling */
        .developer-info {
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
          border-left: 4px solid var(--primary-blue);
          padding: 15px;
          margin-top: 20px;
          border-radius: 8px;
          font-size: 12px;
          color: var(--text-medium);
        }
        
        .developer-info strong {
          color: var(--primary-blue);
        }
        
        /* About page styling */
        .about-container {
          padding: 30px;
          background: linear-gradient(135deg, #ffffff 0%, #fafbfc 100%);
          border-radius: 15px;
          box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
          margin-bottom: 30px;
        }
        
        .about-section {
          margin-bottom: 30px;
          padding-bottom: 20px;
          border-bottom: 1px solid var(--border-gray);
        }
        
        .about-section:last-child {
          border-bottom: none;
        }
        
        .citation-box {
          background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%);
          border-left: 4px solid var(--primary-blue);
          padding: 15px;
          margin: 15px 0;
          border-radius: 8px;
          font-size: 12px;
          line-height: 1.6;
        }
        
        .citation-box code {
          background: white;
          padding: 2px 6px;
          border-radius: 4px;
          font-family: 'Courier New', monospace;
        }
        
        .feature-list {
          list-style-type: none;
          padding-left: 0;
        }
        
        .feature-list li {
          padding: 8px 0;
          padding-left: 30px;
          position: relative;
        }
        
        .feature-list li:before {
          content: '✓';
          position: absolute;
          left: 0;
          color: var(--primary-green);
          font-weight: bold;
        }
      "))
    ),
    sidebarMenu(
      id = "tabs",
      menuItem("Home", tabName = "parallel", icon = icon("chart-line"), selected = TRUE)
    )
  ),
  dashboardBody(
    useShinyjs(),
    # Add app name to browser tab
    tags$head(
      tags$title("FactorGuard - Factor Analysis Tool")
    ),
    div(class = "main-content-container",
        tabItems(
          # Parallel Analysis Tab
          tabItem(tabName = "parallel",
                  fluidRow(
                    column(12,
                           div(class = "box",
                               h2("FactorGuard: A factor-retention decision tool for EFA", style = "text-align: center; margin-bottom: 30px;")
                           )
                    )
                  ),
                  fluidRow(
                    column(4,
                           div(class = "fixed-sidebar",
                               div(class = "box",
                                   h4("Data Input", icon("database")),
                                   fileInput("file_parallel", "Upload Dataset (CSV, Excel, SAV, DTA)", 
                                             accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta"),
                                             width = "100%"),
                                   uiOutput("varSelect_parallel"),
                                   numericInput("n_iter", "Number of Simulations:", value = 20, min = 10, max = 1000),
                                   actionButton("run_parallel", "Run Parallel Analysis", class = "btn-primary", icon = icon("play")),
                                   br(), br(),
                                   div(class = "syntax-help",
                                       h5("Parallel Analysis Notes:", icon("info-circle")),
                                       tags$ul(
                                         tags$li("Parallel analysis compares actual eigenvalues with random data eigenvalues"),
                                         tags$li("Suggested number of factors where actual eigenvalues > random eigenvalues"),
                                         tags$li("More robust than traditional 'eigenvalue > 1' rule"),
                                         tags$li("Scree plot shows eigenvalues in descending order")
                                       )
                                   ),
                                   # Developer information
                                   div(class = "developer-info",
                                       h6("Developer Information", icon("user-circle")),
                                       p("Mudasir Mohammed Ibrahim"),
                                       p(icon("envelope"), " mudassiribrahim30@gmail.com")
                                   )
                               )
                           )
                    ),
                    column(8,
                           div(class = "scrollable-results",
                               tabsetPanel(
                                 tabPanel("Parallel Analysis", 
                                          div(class = "plot-container",
                                              plotOutput("faParallelPlot", height = "500px")
                                          ),
                                          div(class = "download-container",
                                              downloadButton("downloadParallel", "Download Parallel Plot", 
                                                             class = "btn-info btn-block")
                                          ),
                                          verbatimTextOutput("parallelInterpretation")
                                 ),
                                 tabPanel("Scree Plot", 
                                          div(class = "plot-container",
                                              plotlyOutput("screePlot", height = "500px")
                                          )
                                 ),
                                 # NEW ABOUT TAB ADDED HERE
                                 tabPanel("About",
                                          div(class = "about-container",
                                              div(class = "about-section",
                                                  h3("About Me", icon("user")),
                                                  p("My name is ", strong("Mudasir Mohammed Ibrahim"), ", a Registered Nurse with a strong interest in healthcare research and data analysis."),
                                                  p("Through my work in healthcare and research, I identified a recurring challenge: many researchers struggle to apply advanced statistical methods despite their importance for high-quality research. This inspired me to develop FactorGuard—a user-friendly application designed to make factor analysis more accessible, transparent, and practical for researchers across diverse disciplines."),
                                                  p("Contact: ", icon("envelope"), " mudassiribrahim30@gmail.com")
                                              ),
                                              
                                              div(class = "about-section",
                                                  h3("What This App Does", icon("cogs")),
                                                  p("FactorGuard is a Shiny factor analysis visualization tool that helps researchers determine the optimal number of factors to retain in exploratory factor analysis (EFA). The application provides:"),
                                                  tags$ul(class = "feature-list",
                                                          tags$li("Parallel Analysis: Compares actual eigenvalues with eigenvalues from random data to determine factor retention"),
                                                          tags$li("Interactive Scree Plots: Visualize eigenvalues and identify the 'elbow point' in the scree plot"),
                                                          tags$li("Multi-format Data Support: Import data from CSV, Excel, SPSS, and Stata formats"),
                                                          tags$li("Downloadable Results: Export high-quality plots for publication and reports"),
                                                          tags$li("Interactive Visualizations: Hover-enabled plots for detailed examination of eigenvalues")
                                                  ),
                                                  p("The tool is designed to be accessible to both novice and experienced researchers, providing clear interpretations and visual guidance throughout the analysis process.")
                                              ),
                                              
                                              div(class = "about-section",
                                                  h3("Methodological Basis", icon("flask")),
                                                  p("FactorGuard implements ", strong("parallel analysis"), " - a statistically rigorous method for determining the number of factors to retain in factor analysis. The methodology is based on:"),
                                                  p("1. ", strong("Horn's Parallel Analysis (1965):"), " Compares eigenvalues from actual data with eigenvalues from random data matrices. Factors are retained when actual eigenvalues exceed the corresponding percentiles of random data eigenvalues."),
                                                  p("2. ", strong("Comparison with Traditional Methods:"), " Addresses limitations of Kaiser's criterion (eigenvalue > 1) and scree plot interpretation by providing empirical, simulation-based evidence."),
                                                  p("3. ", strong("Monte Carlo Simulations:"), " Generates random data with the same dimensions as the original dataset to create appropriate comparison benchmarks."),
                                                  p("All parallel analysis computations in this application are performed using the ", tags$code("fa.parallel()"), " function from the ", strong("psych"), " package in R. Results have been verified against direct R implementations of the same analyses, producing identical outcomes and ensuring reproducibility and methodological consistency.")
                                              ),
                                              
                                              div(class = "about-section",
                                                  h3("Release Information", icon("calendar")),
                                                  p(strong("Current Version:"), " 1.0.0"),
                                                  p(strong("Initial Release Date:"), " 20th January 2026"),
                                                  p(strong("Last Updated:"), " January 2026"),
                                                  p(strong("Platform:"), " Shiny Web Application (R-based)")
                                              ),
                                              
                                              div(class = "about-section",
                                                  h3("References & Citations", icon("book")),
                                                  p("For methodological details and to cite parallel analysis in your research, please reference:"),
                                                  
                                                  div(class = "citation-box",
                                                      p(strong("Primary Reference on Parallel Analysis:")),
                                                      p("Horn, J. L. (1965). A Rationale and Test for the Number of Factors in Factor Analysis. Psychometrika, 30, 179-185.
https://doi.org/10.1007/BF02289447.")
                                                  ),
                                                  
                                                  div(class = "citation-box",
                                                      p(strong("Application Citation:")),
                                                      p("If you use FactorGuard in your research, please cite as:"),
                                                      p("Ibrahim, M. M. (2026). FactorGuard: A factor-retention decision tool for EFA [Computer software]. Retrieved from [application URL]")

                                                  )
                                              ),
                                              
                                              div(class = "about-section",
                                                  h3("Technical Details", icon("laptop-code")),
                                                  p(strong("Built With:"), " R Shiny, shinydashboard, psych, ggplot2, plotly"),
                                                  p(strong("Source Code:"), " Available on my personal website: https://www.mudasiribrahim.com/"),
                                                  p(strong("License:"), " For research and educational use"),
                                                  p(strong("Compatibility:"), " Works in all modern web browsers with JavaScript enabled")
                                              )
                                          )
                                 )
                               )
                           )
                    )
                  )
          )
        )
    ),
    tags$footer(
      class = "footer",
      div(
        "FactorGuard - Factor Analysis Visualization Tool",
        br(),
        "Developed by ",
        tags$a(href = "mailto:mudassiribrahim30@gmail.com", "Mudasir Mohammed Ibrahim", 
               style = "color: var(--primary-blue); font-weight: 500;"),
        br(),
        "© 2024 All rights reserved"
      )
    )
  )
)

server <- function(input, output, session) {
  # Shared data reactive
  data_shared <- reactive({
    if (!is.null(input$file_parallel)) {
      file <- input$file_parallel
    } else {
      return(NULL)
    }
    
    req(file)
    ext <- tools::file_ext(file$name)
    switch(ext,
           csv = read_csv(file$datapath),
           xlsx = read_excel(file$datapath),
           xls = read_excel(file$datapath),
           sav = read_sav(file$datapath),
           dta = read_dta(file$datapath),
           validate("Unsupported file type")
    )
  })
  
  # Variable selection UI
  output$varSelect_parallel <- renderUI({
    req(data_shared())
    selectInput("vars_parallel", "Select Variables:", choices = names(data_shared()), multiple = TRUE)
  })
  
  # Selected data
  selectedData_parallel <- reactive({
    req(input$vars_parallel)
    data_shared()[, input$vars_parallel, drop = FALSE]
  })
  
  # Reactive values for storing results
  parallel_result <- reactiveVal(NULL)
  
  # Run parallel analysis
  observeEvent(input$run_parallel, {
    req(selectedData_parallel())
    
    # Show notification
    showNotification("Running parallel analysis...", type = "message")
    
    # Run parallel analysis
    tryCatch({
      parallel <- fa.parallel(selectedData_parallel(), fa = "fa", n.iter = input$n_iter, show.legend = TRUE)
      parallel_result(parallel)
      
      # Parallel analysis interpretation
      output$parallelInterpretation <- renderPrint({
        req(parallel_result())
        parallel <- parallel_result()
        cat("PARALLEL ANALYSIS RESULTS\n")
        cat("=========================\n\n")
        cat("Suggested number of factors:", parallel$nfact, "\n")
        cat("INTERPRETATION:\n")
        cat("---------------\n")
        cat("1. Parallel analysis indicates that", parallel$nfact, "factors should be retained for further analysis.\n")
        cat("2. Factors are retained when actual eigenvalues exceed the corresponding\n")
        cat("   percentiles of random data eigenvalues (red line).\n")
        cat("3. This method is more robust than Kaiser's criterion (eigenvalue > 1).\n\n")
        cat("RECOMMENDATIONS:\n")
        cat("----------------\n")
        if (parallel$nfact == 1) {
          cat("- Consider a one-factor solution if conceptually justified.\n")
          cat("- Check scree plot for clear elbow point.\n")
        } else if (parallel$nfact > 1) {
          cat("- Consider extracting", parallel$nfact, "factors.\n")
          cat("- Compare with scree plot results.\n")
          cat("- Evaluate factor interpretability.\n")
        } else {
          cat("- No factors suggested. Data may not be suitable for factor analysis.\n")
          cat("- Check KMO and Bartlett's test if available.\n")
        }
      })
      
      # Parallel analysis plot - with number of factors annotation
      output$faParallelPlot <- renderPlot({
        req(parallel_result())
        parallel <- parallel_result()
        
        # Create custom plot
        n_factors <- length(parallel$fa.values)
        plot_data <- data.frame(
          Factor = 1:n_factors,
          Actual = parallel$fa.values,
          Simulated = parallel$fa.simr
        )
        
        # Determine number of suggested factors
        n_suggested <- parallel$nfact
        
        # Create the plot with annotations
        p <- ggplot(plot_data, aes(x = Factor)) +
          geom_line(aes(y = Actual, color = "Actual Data"), size = 1.5) +
          geom_point(aes(y = Actual, color = "Actual Data"), size = 3) +
          geom_line(aes(y = Simulated, color = "Simulated Data"), size = 1.2, linetype = "dashed") +
          geom_point(aes(y = Simulated, color = "Simulated Data"), size = 2.5, shape = 17) +
          geom_hline(yintercept = 1, linetype = "dotted", color = "darkgray", size = 1) +
          scale_color_manual(values = c("Actual Data" = "#3498db", "Simulated Data" = "#e74c3c")) +
          labs(title = "Parallel Analysis Scree Plot",
               subtitle = paste("Suggested number of factors:", n_suggested),
               x = "Factor Number",
               y = "Eigenvalue",
               color = "Data Type") +
          theme_minimal() +
          theme(
            plot.title = element_text(hjust = 0.5, size = 16, face = "bold"),
            plot.subtitle = element_text(hjust = 0.5, size = 12, color = "#2c3e50"),
            axis.title = element_text(size = 12),
            legend.position = "bottom",
            panel.grid.major = element_line(color = "gray90"),
            panel.grid.minor = element_blank(),
            plot.margin = margin(20, 20, 20, 20)
          ) +
          annotate("text", x = max(plot_data$Factor), y = 1.05, 
                   label = "Eigenvalue = 1", hjust = 1, vjust = 0, color = "darkgray", size = 3.5)
        
        # Add annotation for number of suggested factors
        if (n_suggested > 0 && n_suggested <= nrow(plot_data)) {
          p <- p + 
            annotate("rect", 
                     xmin = n_suggested - 0.3, xmax = n_suggested + 0.3,
                     ymin = plot_data$Actual[n_suggested] - 0.1, 
                     ymax = plot_data$Actual[n_suggested] + 0.1,
                     fill = "#2ecc71", alpha = 0.2) +
            annotate("text", 
                     x = n_suggested, 
                     y = plot_data$Actual[n_suggested] + 0.2,
                     label = paste("Factor", n_suggested),
                     color = "#27ae60", 
                     fontface = "bold",
                     size = 4)
        }
        
        # Add developer watermark
        p <- p + 
          annotate("text", 
                   x = max(plot_data$Factor) * 0.95, 
                   y = max(plot_data$Actual) * 0.05,
                   label = "FactorGuard\nMudasir Mohammed Ibrahim",
                   color = "gray70", 
                   size = 2.5,
                   hjust = 1,
                   alpha = 0.6)
        
        p
      }, height = 500)
      
      # Scree Plot
      output$screePlot <- renderPlotly({
        ev <- eigen(cor(selectedData_parallel(), use = "pairwise.complete.obs"))
        scree_data <- data.frame(
          Factor = 1:length(ev$values), 
          Eigenvalue = ev$values,
          Type = "Actual"
        )
        
        # Add parallel analysis line if available
        if (!is.null(parallel_result())) {
          parallel <- parallel_result()
          if (length(parallel$fa.simr) >= nrow(scree_data)) {
            scree_data$Simulated <- parallel$fa.simr[1:nrow(scree_data)]
          }
        }
        
        # Create base plot
        p <- plot_ly(scree_data, x = ~Factor, y = ~Eigenvalue, 
                     type = 'scatter', mode = 'lines+markers',
                     name = 'Actual Eigenvalues',
                     marker = list(size = 8, color = '#3498db'),
                     line = list(color = '#3498db', width = 3))
        
        # Add simulated line if available
        if ("Simulated" %in% names(scree_data)) {
          p <- p %>% 
            add_trace(x = ~Factor, y = ~Simulated, 
                      type = 'scatter', mode = 'lines+markers',
                      name = 'Parallel Analysis',
                      line = list(color = '#e74c3c', width = 2, dash = 'dash'),
                      marker = list(size = 6, color = '#e74c3c', symbol = 'triangle-up'))
        }
        
        # Add layout with horizontal line annotation
        p <- p %>% 
          layout(
            title = list(
              text = "Scree Plot",
              font = list(size = 16, family = "Arial, sans-serif")
            ),
            xaxis = list(title = "Factor Number", tickmode = "linear"),
            yaxis = list(title = "Eigenvalue"),
            hovermode = 'x unified',
            showlegend = TRUE,
            legend = list(orientation = 'h', y = -0.2),
            shapes = list(
              list(
                type = "line",
                x0 = 0,
                x1 = 1,
                xref = "paper",
                y0 = 1,
                y1 = 1,
                yref = "y",
                line = list(color = "gray", dash = "dot", width = 1)
              )
            ),
            annotations = list(
              list(
                x = 1,
                xref = "paper",
                y = 1.05,
                yref = "y",
                text = "Eigenvalue = 1",
                showarrow = FALSE,
                font = list(color = "darkgray", size = 10),
                xanchor = "right",
                yanchor = "bottom"
              )
            )
          )
        
        p
      })
      
      # Download handler for parallel plot
      output$downloadParallel <- downloadHandler(
        filename = function() { 
          paste0("parallel_analysis_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".png") 
        },
        content = function(file) {
          # Get parallel results
          parallel <- parallel_result()
          n_factors <- length(parallel$fa.values)
          plot_data <- data.frame(
            Factor = 1:n_factors,
            Actual = parallel$fa.values,
            Simulated = parallel$fa.simr
          )
          
          n_suggested <- parallel$nfact
          
          # Create high-quality plot with annotations
          p <- ggplot(plot_data, aes(x = Factor)) +
            geom_line(aes(y = Actual, color = "Actual Data"), size = 1.5) +
            geom_point(aes(y = Actual, color = "Actual Data"), size = 3) +
            geom_line(aes(y = Simulated, color = "Simulated Data"), size = 1.2, linetype = "dashed") +
            geom_point(aes(y = Simulated, color = "Simulated Data"), size = 2.5, shape = 17) +
            geom_hline(yintercept = 1, linetype = "dotted", color = "darkgray", size = 1) +
            scale_color_manual(values = c("Actual Data" = "#3498db", "Simulated Data" = "#e74c3c")) +
            labs(title = "Parallel Analysis Scree Plot",
                 subtitle = paste("Suggested number of factors:", n_suggested),
                 x = "Factor Number",
                 y = "Eigenvalue",
                 color = "Data Type",
                 caption = paste("Generated: ", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), 
                                 "\nFactorGuard")) +
            theme_minimal() +
            theme(
              plot.title = element_text(hjust = 0.5, size = 18, face = "bold"),
              plot.subtitle = element_text(hjust = 0.5, size = 14, color = "#2c3e50"),
              plot.caption = element_text(hjust = 1, size = 9, color = "gray50"),
              axis.title = element_text(size = 12),
              legend.position = "bottom",
              panel.grid.major = element_line(color = "gray90"),
              panel.grid.minor = element_blank(),
              plot.margin = margin(20, 20, 20, 20),
              legend.text = element_text(size = 10)
            ) +
            annotate("text", x = max(plot_data$Factor), y = 1.05, 
                     label = "Eigenvalue = 1", hjust = 1, vjust = 0, color = "darkgray", size = 4)
          
          # Add annotation for number of suggested factors
          if (n_suggested > 0 && n_suggested <= nrow(plot_data)) {
            p <- p + 
              annotate("rect", 
                       xmin = n_suggested - 0.3, xmax = n_suggested + 0.3,
                       ymin = plot_data$Actual[n_suggested] - 0.1, 
                       ymax = plot_data$Actual[n_suggested] + 0.1,
                       fill = "#2ecc71", alpha = 0.2) +
              annotate("text", 
                       x = n_suggested, 
                       y = plot_data$Actual[n_suggested] + 0.25,
                       label = paste("Factor", n_suggested),
                       color = "#27ae60", 
                       fontface = "bold",
                       size = 5)
          }
          
          # Save as high-resolution PNG
          ggsave(file, plot = p, width = 10, height = 7, dpi = 300, bg = "white")
        }
      )
      
      # Download handler for scree plot
      output$downloadScree <- downloadHandler(
        filename = function() { 
          paste0("scree_plot_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".png") 
        },
        content = function(file) {
          ev <- eigen(cor(selectedData_parallel(), use = "pairwise.complete.obs"))
          scree_data <- data.frame(Factor = 1:length(ev$values), Eigenvalue = ev$values)
          
          # Get parallel results if available
          n_suggested <- if (!is.null(parallel_result())) parallel_result()$nfact else NA
          
          p <- ggplot(scree_data, aes(x = Factor, y = Eigenvalue)) +
            geom_point(size = 3, color = "#3498db") + 
            geom_line(color = "#3498db", size = 1.2) +
            geom_hline(yintercept = 1, linetype = "dotted", color = "darkgray", size = 1) +
            labs(title = "Scree Plot", 
                 subtitle = if (!is.na(n_suggested)) paste("Parallel analysis suggests", n_suggested, "factors") else NULL,
                 x = "Factor Number", 
                 y = "Eigenvalue",
                 caption = paste("Generated: ", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), 
                                 "\nFactorGuard")) +
            theme_minimal() +
            theme(
              plot.title = element_text(hjust = 0.5, size = 18, face = "bold"),
              plot.subtitle = element_text(hjust = 0.5, size = 12, color = "#2c3e50"),
              plot.caption = element_text(hjust = 1, size = 9, color = "gray50"),
              axis.title = element_text(size = 12),
              panel.grid.major = element_line(color = "gray90"),
              panel.grid.minor = element_blank(),
              plot.margin = margin(20, 20, 20, 20)
            ) +
            annotate("text", x = max(scree_data$Factor), y = 1.05, 
                     label = "Eigenvalue = 1", hjust = 1, vjust = 0, color = "darkgray", size = 4)
          
          # Add factor annotation if parallel analysis was run
          if (!is.na(n_suggested) && n_suggested > 0 && n_suggested <= nrow(scree_data)) {
            p <- p + 
              annotate("rect", 
                       xmin = n_suggested - 0.3, xmax = n_suggested + 0.3,
                       ymin = scree_data$Eigenvalue[n_suggested] - 0.1, 
                       ymax = scree_data$Eigenvalue[n_suggested] + 0.1,
                       fill = "#2ecc71", alpha = 0.2) +
              annotate("text", 
                       x = n_suggested, 
                       y = scree_data$Eigenvalue[n_suggested] + 0.25,
                       label = paste("Factor", n_suggested),
                       color = "#27ae60", 
                       fontface = "bold",
                       size = 5)
          }
          
          ggsave(file, plot = p, width = 10, height = 7, dpi = 300, bg = "white")
        }
      )
      
      showNotification("Parallel analysis completed!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
}

shinyApp(ui, server)
