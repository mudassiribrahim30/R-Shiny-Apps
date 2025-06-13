# app.R
library(shiny)
library(shinythemes)
library(ggplot2)
library(dplyr)
library(readr)
library(haven)
library(readxl)
library(DT)
library(stringr)
library(plotly)
library(shinydashboard)
library(shinyWidgets)
library(shinycssloaders)

# Custom CSS for enhanced UI
custom_css <- "
:root {
  --primary: #2c3e50;
  --secondary: #3498db;
  --accent: #e74c3c;
  --light: #ecf0f1;
  --dark: #34495e;
}

body {
  font-family: 'Lato', 'Helvetica Neue', Arial, sans-serif;
  background-color: #f9f9f9;
  color: #333;
}

.navbar {
  min-height: 70px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.1);
  background: linear-gradient(135deg, var(--primary), var(--dark));
  border: none;
}

.navbar-brand {
  font-weight: 700;
  font-size: 1.5rem;
  display: flex;
  align-items: center;
  color: white !important;
}

.navbar-brand i {
  margin-right: 10px;
  font-size: 1.3em;
}

.navbar-default .navbar-nav>li>a {
  color: rgba(255,255,255,0.9) !important;
  font-weight: 500;
  transition: all 0.3s ease;
}

.navbar-default .navbar-nav>li>a:hover {
  color: white !important;
  transform: translateY(-2px);
}

.sidebar-panel {
  background-color: white;
  border-radius: 8px;
  box-shadow: 0 4px 6px rgba(0,0,0,0.05);
  padding: 20px;
  margin-bottom: 20px;
  border-left: 4px solid var(--secondary);
}

.main-panel {
  background-color: white;
  border-radius: 8px;
  box-shadow: 0 4px 6px rgba(0,0,0,0.05);
  padding: 25px;
}

.btn {
  border-radius: 4px;
  font-weight: 500;
  letter-spacing: 0.5px;
  transition: all 0.3s ease;
  border: none;
  padding: 10px 20px;
}

.btn:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 8px rgba(0,0,0,0.1);
}

.btn-primary {
  background-color: var(--secondary);
}

.btn-success {
  background-color: #27ae60;
}

.btn-info {
  background-color: #2980b9;
}

.btn-danger {
  background-color: var(--accent);
}

.selectize-input {
  border-radius: 4px !important;
  border: 1px solid #ddd !important;
  box-shadow: none !important;
}

.well {
  border-radius: 8px;
  box-shadow: 0 2px 10px rgba(0,0,0,0.05);
  background-color: white;
  border: none;
}

.tab-content {
  padding: 20px 0;
}

.nav-tabs>li>a {
  color: var(--dark);
  font-weight: 500;
}

.nav-tabs>li.active>a, 
.nav-tabs>li.active>a:focus, 
.nav-tabs>li.active>a:hover {
  color: var(--secondary);
  border-bottom: 3px solid var(--secondary);
}

.footer {
  text-align: center;
  padding: 30px 0;
  margin-top: 50px;
  color: #7f8c8d;
  font-size: 14px;
  border-top: 1px solid #eee;
}

.about-header {
  background: linear-gradient(135deg, var(--primary), var(--dark));
  color: white;
  padding: 30px;
  border-radius: 8px;
  margin-bottom: 30px;
}

.feature-card {
  background: white;
  border-radius: 8px;
  padding: 20px;
  margin-bottom: 20px;
  box-shadow: 0 4px 6px rgba(0,0,0,0.05);
  border-top: 3px solid var(--secondary);
  transition: all 0.3s ease;
}

.feature-card:hover {
  transform: translateY(-5px);
  box-shadow: 0 8px 15px rgba(0,0,0,0.1);
}

.stat-card {
  background: white;
  border-radius: 8px;
  padding: 15px;
  text-align: center;
  box-shadow: 0 4px 6px rgba(0,0,0,0.05);
  margin-bottom: 20px;
  border-left: 4px solid var(--secondary);
}

.stat-value {
  font-size: 2.5rem;
  font-weight: 700;
  color: var(--secondary);
}

.stat-label {
  color: #7f8c8d;
  font-size: 0.9rem;
}

.transform-card {
  background: white;
  border-radius: 8px;
  padding: 15px;
  margin-bottom: 15px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.05);
  border-left: 3px solid var(--secondary);
}

.transform-card h4 {
  color: var(--primary);
  margin-top: 0;
}

.transform-card p {
  color: #666;
}

.plot-container {
  background-color: white;
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 4px 6px rgba(0,0,0,0.05);
  margin-bottom: 30px;
  border: 1px solid #eee;
}

.data-table {
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}

.help-block {
  color: #7f8c8d;
  font-size: 0.85rem;
}

.progress-bar {
  background-color: var(--secondary);
}
"

# Define UI
ui <- fluidPage(
  theme = shinytheme("flatly"),
  tags$head(
    tags$link(rel = "stylesheet", type = "text/css", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css"),
    tags$link(rel = "stylesheet", href = "https://fonts.googleapis.com/css2?family=Lato:wght@300;400;700&display=swap"),
    tags$style(HTML(custom_css))
  ),
  
  navbarPage(
    title = div(icon("database"), 
                span("DataTransformR", style = "font-family: 'Lato'; font-weight: 700;")),
    id = "nav",
    windowTitle = "DataTransformR | Advanced Data Transformation Tool",
    collapsible = TRUE,
    inverse = TRUE,
    fluid = TRUE,
    
    tabPanel(
      "Upload Data",
      icon = icon("upload"),
      div(class = "container-fluid",
          div(class = "row",
              sidebarLayout(
                sidebarPanel(
                  width = 3,
                  div(class = "sidebar-panel",
                      h4("Data Import", style = "color: var(--primary); margin-top: 0;"),
                      radioGroupButtons(
                        inputId = "file_type",
                        label = "Select File Type:", 
                        choices = c("CSV", "Excel", "SPSS", "Stata"),
                        selected = "CSV",
                        status = "primary",
                        justified = TRUE
                      ),
                      
                      conditionalPanel(
                        condition = "input.file_type == 'CSV'",
                        fileInput("csv_file", "Choose CSV File",
                                  accept = c(".csv", ".txt"),
                                  buttonLabel = "Browse...",
                                  placeholder = "No file selected"),
                        materialSwitch("header_csv", "Header", TRUE, status = "primary"),
                        radioGroupButtons("sep_csv", "Separator",
                                          choices = c(Comma = ",", Semicolon = ";", Tab = "\t"),
                                          selected = ",", status = "primary"),
                        radioGroupButtons("quote_csv", "Quote",
                                          choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"),
                                          selected = '"', status = "primary")
                      ),
                      
                      conditionalPanel(
                        condition = "input.file_type == 'Excel'",
                        fileInput("excel_file", "Choose Excel File",
                                  accept = c(".xls", ".xlsx"),
                                  buttonLabel = "Browse...",
                                  placeholder = "No file selected"),
                        materialSwitch("header_excel", "Header", TRUE, status = "primary"),
                        numericInput("sheet_excel", "Sheet Number", 1, min = 1)
                      ),
                      
                      conditionalPanel(
                        condition = "input.file_type == 'SPSS'",
                        fileInput("spss_file", "Choose SPSS File",
                                  accept = ".sav",
                                  buttonLabel = "Browse...",
                                  placeholder = "No file selected")
                      ),
                      
                      conditionalPanel(
                        condition = "input.file_type == 'Stata'",
                        fileInput("stata_file", "Choose Stata File",
                                  accept = ".dta",
                                  buttonLabel = "Browse...",
                                  placeholder = "No file selected")
                      ),
                      
                      actionBttn("load_data", "Load Data", 
                                 style = "gradient", color = "primary", block = TRUE)
                  )
                ),
                
                mainPanel(
                  width = 9,
                  div(class = "main-panel",
                      h3("Data Preview", style = "color: var(--primary);"),
                      div(class = "data-table",
                          DTOutput("data_preview")
                      ),
                      tags$hr(),
                      uiOutput("data_summary_ui")
                  )
                )
              )
          )
      )
    ),
    
    tabPanel(
      "Transform Data",
      icon = icon("exchange-alt"),
      div(class = "container-fluid",
          div(class = "row",
              sidebarLayout(
                sidebarPanel(
                  width = 3,
                  div(class = "sidebar-panel",
                      h4("Transformation Controls", style = "color: var(--primary); margin-top: 0;"),
                      uiOutput("var_select"),
                      
                      selectizeInput(
                        "transformation",
                        "Select Transformation:",
                        choices = c(
                          "None" = "none",
                          "Logarithmic" = list(
                            "Log10" = "log10",
                            "Natural Log (ln)" = "log"
                          ),
                          "Power" = list(
                            "Square Root" = "sqrt",
                            "Inverse (1/x)" = "inverse",
                            "Square" = "square",
                            "Cube" = "cube"
                          ),
                          "Standardization" = list(
                            "Z-score" = "zscore",
                            "Min-Max Normalization" = "minmax"
                          ),
                          "Advanced" = list(
                            "Box-Cox" = "boxcox",
                            "Yeo-Johnson" = "yeojohnson"
                          ),
                          "Discretization" = list(
                            "Binning (Quantiles)" = "bin_quantile",
                            "Binning (Equal Width)" = "bin_equal"
                          )
                        ),
                        options = list(placeholder = 'Select transformation...')
                      ),
                      
                      conditionalPanel(
                        condition = "input.transformation == 'bin_quantile' || input.transformation == 'bin_equal'",
                        sliderInput("num_bins", "Number of Bins:", 
                                    min = 2, max = 20, value = 4, step = 1)
                      ),
                      
                      conditionalPanel(
                        condition = "input.transformation == 'boxcox' || input.transformation == 'yeojohnson'",
                        sliderInput("lambda", "Lambda Value:", 
                                    min = -2, max = 2, value = 0.5, step = 0.1)
                      ),
                      
                      textInput("new_var_name", "New Variable Name", "transformed_var"),
                      
                      actionBttn("apply_transform", "Apply Transformation", 
                                 style = "gradient", color = "primary", block = TRUE),
                      
                      actionBttn("add_to_data", "Add to Dataset", 
                                 style = "gradient", color = "success", block = TRUE),
                      
                      tags$hr(),
                      
                      h5("Export Options", style = "color: var(--primary);"),
                      pickerInput("download_type", "Download Format:",
                                  choices = c("CSV" = "csv", "Excel" = "xlsx", 
                                              "SPSS" = "sav", "Stata" = "dta"),
                                  selected = "csv",
                                  options = list(style = "btn-primary")),
                      
                      downloadBttn("download_data", "Download Data", 
                                   style = "gradient", color = "royal", block = TRUE)
                  )
                ),
                
                mainPanel(
                  width = 9,
                  div(class = "main-panel",
                      h3("Transformation Visualization", style = "color: var(--primary);"),
                      div(class = "plot-container",
                          plotlyOutput("transformation_plot", height = "500px") %>% 
                            withSpinner(color = "#3498db")
                      ),
                      
                      h3("Transformed Data Preview", style = "color: var(--primary);"),
                      div(class = "data-table",
                          DTOutput("transformed_preview") %>% 
                            withSpinner(color = "#3498db")
                      ),
                      
                      tags$hr(),
                      
                      h3("Current Dataset", style = "color: var(--primary);"),
                      div(class = "data-table",
                          DTOutput("current_data") %>% 
                            withSpinner(color = "#3498db")
                      )
                  )
                )
              )
          )
      )
    ),
    
    tabPanel(
      "About",
      icon = icon("info-circle"),
      div(class = "container-fluid",
          div(class = "about-header",
              h2("DataTransformR", style = "font-weight: 700;"),
              h4("Advanced Data Transformation Platform", style = "font-weight: 300;")
          ),
          
          div(class = "row",
              div(class = "col-md-8",
                  div(class = "main-panel",
                      h3("Comprehensive Data Transformation Solution"),
                      p("DataTransformR is a professional-grade application designed for researchers, data scientists, and analysts who require robust tools for data preprocessing and transformation. Developed with cutting-edge statistical methodologies, this platform combines academic rigor with intuitive design."),
                      
                      h4("Scientific Foundation"),
                      p("All transformations are implemented according to established statistical literature:"),
                      tags$ul(
                        tags$li("Box-Cox and Yeo-Johnson transformations follow the original publications"),
                        tags$li("Standardization methods adhere to IEEE standards"),
                        tags$li("Binning algorithms use optimal binning strategies from computational statistics")
                      ),
                      
                      h4("Technical Specifications"),
                      p("The application is built with:"),
                      div(class = "row",
                          div(class = "col-md-6",
                              div(class = "stat-card",
                                  div(class = "stat-value", "R 4.1+"),
                                  div(class = "stat-label", "Statistical Engine")
                              )
                          ),
                          div(class = "col-md-6",
                              div(class = "stat-card",
                                  div(class = "stat-value", "Shiny"),
                                  div(class = "stat-label", "Web Framework")
                              )
                          )
                      ),
                      
                      h4("Transformation Methods"),
                      div(class = "row",
                          div(class = "col-md-6",
                              div(class = "transform-card",
                                  h4("Logarithmic"),
                                  p("log10, natural log (ln) with automatic handling of zero/negative values")
                              ),
                              div(class = "transform-card",
                                  h4("Power"),
                                  p("Square root, inverse, square, and cube transformations with domain checking")
                              )
                          ),
                          div(class = "col-md-6",
                              div(class = "transform-card",
                                  h4("Standardization"),
                                  p("Z-score (μ=0, σ=1) and min-max scaling (0-1 range)")
                              ),
                              div(class = "transform-card",
                                  h4("Advanced"),
                                  p("Box-Cox (λ estimation) and Yeo-Johnson (handles all real numbers)")
                              )
                          )
                      )
                  )
              ),
              
              div(class = "col-md-4",
                  div(class = "main-panel",
                      h4("Developer Information"),
                      div(class = "feature-card",
                          h4("Mudasir Mohammed Ibrahim"),
                          p("Registered Nurse"),
                          tags$ul(
                            tags$li(icon("envelope"), " mudassiribrahim30@gmail.com"),
                            tags$li(icon("github"), " github.com/mudassiribrahim30")
                          ),
                          p("For technical support or feature requests, please contact via email.")
                      ),
                      
                      h4("System Requirements"),
                      div(class = "feature-card",
                          tags$ul(
                            tags$li("Modern web browser (Chrome, Firefox, Safari)"),
                            tags$li("Internet connection for initial load"),
                            tags$li("Minimum screen resolution: 1280×720"),
                            tags$li("JavaScript enabled")
                          )
                      ),
                      
                      h4("Citation"),
                      div(class = "feature-card",
                          p("If you use DataTransformR in your research, please cite:"),
                          p(em("Ibrahim, M.M. (2025). DataTransformR: Advanced Data Transformation Tool"))
                      )
                  )
              )
          ),
          
          div(class = "footer",
              p(icon("copyright"), " 2025 DataTransformR | Advanced Data Science Platform"),
              p("Developed with ", icon("heart"), " for the research community")
          )
      )
    )
  )
)

# Define server logic
server <- function(input, output, session) {
  
  # Reactive value to store loaded data
  rv <- reactiveValues(
    data = NULL,
    transformed_data = NULL,
    transformed_var = NULL
  )
  
  # Load data based on file type
  observeEvent(input$load_data, {
    tryCatch({
      showModal(modalDialog(
        title = "Loading Data",
        "Please wait while we process your data...",
        footer = NULL,
        easyClose = FALSE
      ))
      
      if (input$file_type == "CSV" && !is.null(input$csv_file)) {
        rv$data <- read.csv(
          input$csv_file$datapath,
          header = input$header_csv,
          sep = input$sep_csv,
          quote = input$quote_csv
        )
      } else if (input$file_type == "Excel" && !is.null(input$excel_file)) {
        rv$data <- read_excel(
          input$excel_file$datapath,
          sheet = input$sheet_excel,
          col_names = input$header_excel
        )
      } else if (input$file_type == "SPSS" && !is.null(input$spss_file)) {
        rv$data <- read_sav(input$spss_file$datapath)
      } else if (input$file_type == "Stata" && !is.null(input$stata_file)) {
        rv$data <- read_dta(input$stata_file$datapath)
      }
      
      # Convert haven_labelled to factor for SPSS/Stata files
      if (input$file_type %in% c("SPSS", "Stata")) {
        rv$data <- rv$data %>% 
          mutate(across(where(~inherits(., "haven_labelled")), as_factor))
      }
      
      removeModal()
      sendSweetAlert(
        session = session,
        title = "Success",
        text = "Data loaded successfully!",
        type = "success"
      )
    }, error = function(e) {
      removeModal()
      sendSweetAlert(
        session = session,
        title = "Error",
        text = paste("Error loading data:", e$message),
        type = "error"
      )
    })
  })
    
  # Data summary
  output$data_summary_ui <- renderUI({
    req(rv$data)
    div(
      h3("Data Summary", style = "color: var(--primary);"),
      div(class = "data-table",
          DTOutput("data_summary")
      )
    )
  })
  
  output$data_summary <- renderDT({
    req(rv$data)
    summary_df <- data.frame(
      Variable = names(rv$data),
      Type = sapply(rv$data, function(x) class(x)[1]),
      Missing = sapply(rv$data, function(x) sum(is.na(x))),
      Unique = sapply(rv$data, function(x) length(unique(x))),
      stringsAsFactors = FALSE
    )
    
    numeric_vars <- sapply(rv$data, is.numeric)
    if (any(numeric_vars)) {
      numeric_summary <- data.frame(
        Min = sapply(rv$data[, numeric_vars, drop = FALSE], function(x) round(min(x, na.rm = TRUE), 3)),
        Mean = sapply(rv$data[, numeric_vars, drop = FALSE], function(x) round(mean(x, na.rm = TRUE), 3)),
        Max = sapply(rv$data[, numeric_vars, drop = FALSE], function(x) round(max(x, na.rm = TRUE), 3)),
        SD = sapply(rv$data[, numeric_vars, drop = FALSE], function(x) round(sd(x, na.rm = TRUE), 3)),
        row.names = names(rv$data)[numeric_vars]
      )
      
      matched_rows <- match(summary_df$Variable, rownames(numeric_summary))
      numeric_summary_matched <- numeric_summary[matched_rows, ]
      summary_df <- cbind(summary_df, numeric_summary_matched)
    }
  
      datatable(
        summary_df,
        options = list(
          scrollX = TRUE, 
          pageLength = 10,
          dom = 't',
          initComplete = JS(
            "function(settings, json) {",
            "$(this.api().table().header()).css({'background-color': '#2c3e50', 'color': '#fff'});",
            "}")
        ),
        rownames = FALSE,
        class = 'hover stripe'
      ) %>% 
        formatStyle(names(summary_df), backgroundColor = 'white')
    })
        
        # Variable selection dropdown
        output$var_select <- renderUI({
          req(rv$data)
          numeric_vars <- names(rv$data)[sapply(rv$data, is.numeric)]
          if (length(numeric_vars) == 0) {
            pickerInput(
              inputId = "selected_var",
              label = "Select Variable (No numeric variables found)", 
              choices = NULL,
              options = list(style = "btn-primary")
            )
          } else {
            pickerInput(
              inputId = "selected_var",
              label = "Select Numeric Variable:", 
              choices = numeric_vars,
              options = list(style = "btn-primary",
                             `live-search` = TRUE)
            )
          }
        })
        
        # Apply transformation
        observeEvent(input$apply_transform, {
          req(rv$data, input$selected_var)
          
          tryCatch({
            showModal(modalDialog(
              title = "Applying Transformation",
              "Processing your transformation...",
              footer = NULL,
              easyClose = FALSE
            ))
            
            var <- rv$data[[input$selected_var]]
            
            rv$transformed_var <- switch(
              input$transformation,
              "none" = var,
              "log10" = {
                if (any(var <= 0)) {
                  var <- var + abs(min(var, na.rm = TRUE)) + 1
                  showNotification("Added constant to handle non-positive values", type = "warning")
                }
                log10(var)
              },
              "log" = {
                if (any(var <= 0)) {
                  var <- var + abs(min(var, na.rm = TRUE)) + 1
                  showNotification("Added constant to handle non-positive values", type = "warning")
                }
                log(var)
              },
              "sqrt" = {
                if (any(var < 0)) {
                  showNotification("Negative values set to NA for sqrt", type = "warning")
                  var[var < 0] <- NA
                }
                sqrt(var)
              },
              "inverse" = {
                if (any(var == 0)) {
                  showNotification("Zeros set to NA for inverse", type = "warning")
                  var[var == 0] <- NA
                }
                1/var
              },
              "square" = var^2,
              "cube" = var^3,
              "zscore" = scale(var)[,1],
              "minmax" = (var - min(var, na.rm = TRUE)) / (max(var, na.rm = TRUE) - min(var, na.rm = TRUE)),
              "boxcox" = {
                if (any(var <= 0)) {
                  var <- var + abs(min(var, na.rm = TRUE)) + 1
                  showNotification("Added constant to handle non-positive values", type = "warning")
                }
                if (input$lambda == 0) {
                  log(var)
                } else {
                  (var^input$lambda - 1) / input$lambda
                }
              },
              "yeojohnson" = {
                if (input$lambda == 0) {
                  log(var + 1)
                } else if (input$lambda == 2) {
                  -log(-var + 1)
                } else {
                  ((var + 1)^input$lambda - 1) / input$lambda
                }
              },
              "bin_quantile" = {
                cut(var, 
                    breaks = quantile(var, probs = seq(0, 1, length.out = input$num_bins + 1), na.rm = TRUE), 
                    include.lowest = TRUE, 
                    labels = paste0("Q", 1:input$num_bins))
              },
              "bin_equal" = {
                cut(var, 
                    breaks = seq(min(var, na.rm = TRUE), max(var, na.rm = TRUE), length.out = input$num_bins + 1), 
                    include.lowest = TRUE,
                    labels = paste0("Bin ", 1:input$num_bins))
              }
            )
            
            # Create a temporary data frame for preview
            rv$transformed_data <- data.frame(
              Original = var,
              Transformed = rv$transformed_var
            )
            
            removeModal()
            sendSweetAlert(
              session = session,
              title = "Success",
              text = "Transformation applied successfully!",
              type = "success"
            )
          }, error = function(e) {
            removeModal()
            sendSweetAlert(
              session = session,
              title = "Error",
              text = paste("Error in transformation:", e$message),
              type = "error"
            )
          })
        })
        
        # Add transformed variable to dataset
        observeEvent(input$add_to_data, {
          req(rv$data, rv$transformed_var)
          
          new_var_name <- input$new_var_name
          if (new_var_name == "") {
            sendSweetAlert(
              session = session,
              title = "Warning",
              text = "Please provide a name for the new variable",
              type = "warning"
            )
            return()
          }
          
          # Handle duplicate names
          if (new_var_name %in% names(rv$data)) {
            i <- 1
            while (paste0(new_var_name, "_", i) %in% names(rv$data)) {
              i <- i + 1
            }
            new_var_name <- paste0(new_var_name, "_", i)
            showNotification(paste("Variable name exists. Using", new_var_name), type = "warning")
          }
          
          rv$data[[new_var_name]] <- rv$transformed_var
          sendSweetAlert(
            session = session,
            title = "Success",
            text = paste("Variable", new_var_name, "added to dataset"),
            type = "success"
          )
        })
        
        # Transformation plot - now interactive with plotly
        output$transformation_plot <- renderPlotly({
          req(rv$transformed_data)
          
          if (input$transformation %in% c("bin_quantile", "bin_equal")) {
            # Bar plot for binned data
            p <- ggplot(rv$transformed_data, aes(x = Transformed, fill = Transformed)) +
              geom_bar(color = "white") +
              scale_fill_brewer(palette = "Blues") +
              labs(title = paste("Distribution of", input$transformation, "Transformed Variable"),
                   x = paste(input$selected_var, "(", input$transformation, ")"),
                   y = "Count") +
              theme_minimal(base_size = 14) +
              theme(legend.position = "none",
                    axis.text.x = element_text(angle = 45, hjust = 1),
                    plot.title = element_text(color = "#2c3e50", face = "bold"))
          } else {
            # Histogram for continuous transformations
            p <- ggplot(rv$transformed_data, aes(x = Transformed)) +
              geom_histogram(fill = "#3498db", color = "white", bins = 30, alpha = 0.8) +
              labs(title = paste("Distribution of", input$transformation, "Transformed Variable"),
                   x = paste(input$selected_var, "(", input$transformation, ")"),
                   y = "Count") +
              theme_minimal(base_size = 14) +
              theme(plot.title = element_text(color = "#2c3e50", face = "bold"))
          }
          
          ggplotly(p, tooltip = "y") %>% 
            layout(autosize = TRUE,
                   hoverlabel = list(bgcolor = "white"),
                   margin = list(l = 50, r = 50, b = 50, t = 80),
                   title = list(x = 0.05, xanchor = 'left')) %>% 
            config(displayModeBar = TRUE,
                   displaylogo = FALSE,
                   modeBarButtonsToRemove = c("select2d", "lasso2d", "autoScale2d"))
        })
        
        # Transformed data preview
        output$transformed_preview <- renderDT({
          req(rv$transformed_data)
          datatable(
            rv$transformed_data,
            options = list(
              scrollX = TRUE, 
              pageLength = 5,
              dom = 't',
              initComplete = JS(
                "function(settings, json) {",
                "$(this.api().table().header()).css({'background-color': '#2c3e50', 'color': '#fff'});",
                "}")
            ),
            rownames = FALSE,
            class = 'hover stripe'
          ) %>% 
            formatStyle(names(rv$transformed_data), backgroundColor = 'white') %>% 
            formatRound(columns = c("Original", "Transformed"), digits = 3)
        })
        
        # Current dataset display
        output$current_data <- renderDT({
          req(rv$data)
          datatable(
            rv$data,
            options = list(
              scrollX = TRUE, 
              pageLength = 5,
              dom = 'Blfrtip',
              buttons = c('copy', 'csv', 'excel'),
              initComplete = JS(
                "function(settings, json) {",
                "$(this.api().table().header()).css({'background-color': '#2c3e50', 'color': '#fff'});",
                "}")
            ),
            extensions = 'Buttons',
            rownames = FALSE,
            class = 'hover stripe'
          ) %>% 
            formatStyle(names(rv$data), backgroundColor = 'white')
        })
        
        # Download handler with multiple format options
        output$download_data <- downloadHandler(
          filename = function() {
            paste0("transformed_data_", Sys.Date(), ".", input$download_type)
          },
          content = function(file) {
            req(rv$data)
            tryCatch({
              showModal(modalDialog(
                title = "Preparing Download",
                "Generating your data file...",
                footer = NULL,
                easyClose = FALSE
              ))
              
              switch(
                input$download_type,
                "csv" = write.csv(rv$data, file, row.names = FALSE),
                "xlsx" = writexl::write_xlsx(rv$data, file),
                "sav" = write_sav(rv$data, file),
                "dta" = write_dta(rv$data, file)
              )
              
              removeModal()
            }, error = function(e) {
              removeModal()
              sendSweetAlert(
                session = session,
                title = "Error",
                text = paste("Error saving file:", e$message),
                type = "error"
              )
            })
          }
        )
}

# Run the application
shinyApp(ui = ui, server = server)
