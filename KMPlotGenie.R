library(shiny)
library(shinythemes)
library(survival)
library(survminer)
library(haven)
library(readxl)
library(DT)
library(officer)
library(flextable)
library(bslib)
library(shinyWidgets)
library(ggplot2)
library(ggiraph)
library(colourpicker)
library(dplyr)
library(tools)
library(paletteer)
library(shinycssloaders)

# Custom CSS for eye-friendly styling
custom_css <- "
:root {
  --primary: #2c7a4d;
  --secondary: #4a9d63;
  --accent: #e67e22;
  --light: #f5f5f5;
  --dark: #333333;
}

body {
  background-color: #f5f5f5;
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}

.navbar {
  background-color: var(--primary) !important;
  box-shadow: 0 4px 20px 0 rgba(0,0,0,0.1);
  border: none !important;
}

.navbar-brand {
  font-weight: 700;
  font-size: 1.8rem !important;
  color: white !important;
}

.tab-content {
  background-color: white;
  border-radius: 8px;
  box-shadow: 0 2px 10px rgba(0,0,0,0.05);
  padding: 20px;
  margin-top: 10px;
}

.card {
  border: none !important;
  border-radius: 8px !important;
  box-shadow: 0 2px 8px rgba(0,0,0,0.05);
  margin-bottom: 20px;
}

.card-header {
  background-color: var(--primary) !important;
  color: white !important;
  border-radius: 8px 8px 0 0 !important;
  font-weight: 600;
}

.btn-primary {
  background-color: var(--primary) !important;
  border: none !important;
}

.btn-primary:hover {
  background-color: var(--secondary) !important;
}

.plot-container {
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}

.developer-info {
  position: fixed;
  bottom: 10px;
  right: 10px;
  background: rgba(255,255,255,0.9);
  padding: 8px 15px;
  border-radius: 20px;
  box-shadow: 0 2px 10px rgba(0,0,0,0.1);
  font-size: 0.85rem;
  color: var(--dark);
}

.developer-info a {
  color: var(--secondary) !important;
  font-weight: 600;
}

.well {
  background-color: white !important;
  border: 1px solid #e0e0e0 !important;
}
"

ui <- fluidPage(
  theme = bslib::bs_theme(
    version = 5,
    primary = "#2c7a4d",
    secondary = "#4a9d63",
    success = "#4CAF50",
    info = "#00bcd4",
    warning = "#e67e22",
    danger = "#e74c3c",
    base_font = font_google("Roboto"),
    heading_font = font_google("Montserrat"),
    bg = "#f5f5f5",
    fg = "#333333"
  ),
  
  tags$head(
    tags$style(HTML(custom_css)),
    tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css")
  ),
  
  navbarPage(
    title = div(
      span("KMPlotGenie", style = "font-weight: 700; font-size: 1.8rem;"),
      span(" v1.0", style = "font-size: 0.8rem; vertical-align: super;")
    ),
    id = "navbar",
    windowTitle = "KMPlotGenie - Interactive Survival Analysis",
    collapsible = TRUE,
    inverse = TRUE,
    
    tabPanel(
      title = span(icon("database"), "Data Upload"),
      div(
        class = "tab-content",
        h3("Upload Your Dataset", class = "text-center mb-4"),
        fluidRow(
          column(
            width = 6,
            div(
              class = "card",
              div(
                class = "card-header",
                "File Upload"
              ),
              div(
                class = "card-body",
                fileInput(
                  "file",
                  "Choose file (CSV, Excel, Stata, SPSS):",
                  multiple = FALSE,
                  accept = c(".csv", ".xls", ".xlsx", ".dta", ".sav"),
                  buttonLabel = "Browse...",
                  placeholder = "No file selected"
                ),
                
                conditionalPanel(
                  condition = "output.file_type == 'csv'",
                  awesomeCheckbox(
                    inputId = "header",
                    label = "Header",
                    value = TRUE,
                    status = "primary"
                  ),
                  radioGroupButtons(
                    inputId = "sep",
                    label = "Separator:",
                    choices = c(
                      "Comma" = ",",
                      "Semicolon" = ";",
                      "Tab" = "\t"
                    ),
                    selected = ",",
                    status = "primary",
                    size = "sm"
                  ),
                  radioGroupButtons(
                    inputId = "quote",
                    label = "Quote:",
                    choices = c(
                      "None" = "",
                      "Double Quote" = '"',
                      "Single Quote" = "'"
                    ),
                    selected = '"',
                    status = "primary",
                    size = "sm"
                  )
                ),
                
                actionBttn(
                  inputId = "load_data",
                  label = "Load Data",
                  style = "unite",
                  color = "primary",
                  block = TRUE,
                  size = "md"
                )
              )
            )
          ),
          column(
            width = 6,
            div(
              class = "card",
              div(
                class = "card-header",
                "Data Preview"
              ),
              div(
                class = "card-body",
                DTOutput("data_preview") %>% withSpinner(color = "#2c7a4d")
              )
            )
          )
        )
      )
    ),
    
    tabPanel(
      title = span(icon("chart-line"), "Survival Analysis"),
      div(
        class = "tab-content",
        h3("Kaplan-Meier Survival Analysis", class = "text-center mb-4"),
        fluidRow(
          column(
            width = 4,
            div(
              class = "card",
              div(
                class = "card-header",
                "Analysis Parameters"
              ),
              div(
                class = "card-body",
                uiOutput("time_var_selector"),
                uiOutput("event_var_selector"),
                uiOutput("group_var_selector"),
                awesomeCheckbox(
                  inputId = "conf_int",
                  label = "Show confidence interval",
                  value = TRUE,
                  status = "primary"
                ),
                awesomeCheckbox(
                  inputId = "show_pval",
                  label = "Show p-value",
                  value = TRUE,
                  status = "primary"
                ),
                awesomeCheckbox(
                  inputId = "risk_table",
                  label = "Show risk table",
                  value = TRUE,
                  status = "primary"
                ),
                actionBttn(
                  inputId = "generate_plot",
                  label = "Generate Survival Plot",
                  style = "unite",
                  color = "primary",
                  block = TRUE,
                  size = "md"
                )
              )
            ),
            div(
              class = "card mt-4",
              div(
                class = "card-header",
                "Log-Rank Test"
              ),
              div(
                class = "card-body",
                verbatimTextOutput("logrank_test"),
                downloadButton("download_logrank", "Download Results", class = "btn-primary btn-block")
              )
            )
          ),
          column(
            width = 8,
            div(
              class = "card",
              div(
                class = "card-header d-flex justify-content-between align-items-center",
                span("Kaplan-Meier Plot"),
                dropdownButton(
                  circle = TRUE,
                  status = "primary",
                  icon = icon("gear"),
                  width = "300px",
                  tooltip = tooltipOptions(title = "Plot customization"),
                  div(
                    class = "plot-customization",
                    h4("Plot Customization"),
                    textInput("plot_title", "Plot title", value = "Kaplan-Meier Survival Curve"),
                    textInput("xlab", "X-axis label", value = "Time"),
                    textInput("ylab", "Y-axis label", value = "Survival Probability"),
                    textInput("legend_title", "Legend title", value = "Group"),
                    numericInput("font_size", "Font size", value = 12, min = 8, max = 20),
                    selectInput(
                      "theme",
                      "Plot theme",
                      choices = c(
                        "Minimal" = "minimal",
                        "Classic" = "classic",
                        "Gray" = "gray",
                        "BW" = "bw",
                        "Light" = "light",
                        "Dark" = "dark"
                      ),
                      selected = "minimal"
                    ),
                    awesomeCheckbox(
                      "median_line",
                      "Show median survival line",
                      value = TRUE,
                      status = "primary"
                    ),
                    awesomeCheckbox(
                      "cumulative_events",
                      "Show cumulative events",
                      value = FALSE,
                      status = "primary"
                    ),
                    awesomeCheckbox(
                      "censor_marks",
                      "Show censor marks",
                      value = TRUE,
                      status = "primary"
                    ),
                    uiOutput("color_pickers")
                  )
                )
              ),
              div(
                class = "card-body",
                girafeOutput("surv_plot", width = "100%", height = "600px") %>% 
                  withSpinner(color = "#2c7a4d", type = 6),
                div(
                  class = "d-flex justify-content-between mt-3",
                  downloadButton("download_plot_png", "Download as PNG", class = "btn-primary"),
                  downloadButton("download_plot_pdf", "Download as PDF", class = "btn-primary"),
                  downloadButton("download_plot_svg", "Download as SVG", class = "btn-primary")
                )
              )
            )
          )
        )
      )
    ),
    
    tabPanel(
      title = span(icon("table"), "Survival Table"),
      div(
        class = "tab-content",
        h3("Survival Table Generator", class = "text-center mb-4"),
        fluidRow(
          column(
            width = 4,
            div(
              class = "card",
              div(
                class = "card-header",
                "Table Parameters"
              ),
              div(
                class = "card-body",
                textInput(
                  "time_points",
                  "Time points (comma separated):",
                  value = "",
                  placeholder = "e.g., 30, 60, 90, 180, 365"
                ),
                awesomeCheckbox(
                  "show_ci_table",
                  "Show confidence intervals",
                  value = TRUE,
                  status = "primary"
                ),
                awesomeCheckbox(
                  "show_median",
                  "Show median survival",
                  value = TRUE,
                  status = "primary"
                ),
                actionBttn(
                  "generate_table",
                  "Generate Survival Table",
                  style = "unite",
                  color = "primary",
                  block = TRUE,
                  size = "md"
                )
              )
            )
          ),
          column(
            width = 8,
            div(
              class = "card",
              div(
                class = "card-header",
                "Survival Table"
              ),
              div(
                class = "card-body",
                DTOutput("surv_table") %>% withSpinner(color = "#2c7a4d"),
                downloadButton("download_table_docx", "Download as Word", class = "btn-primary mt-3")
              )
            )
          )
        )
      )
    )
  ),
  
  div(
    class = "developer-info",
    "Developed by ",
    tags$b("Mudasir Mohammed Ibrahim"),
    br(),
    "For suggestions: ",
    tags$a(href = "mailto:mudassiribrahim30@gmail.com", "mudassiribrahim30@gmail.com")
  )
)

server <- function(input, output, session) {
  
  rv <- reactiveValues(
    data = NULL,
    fit = NULL,
    surv_table = NULL,
    logrank = NULL,
    group_levels = NULL,
    file_type = NULL,
    plot_initialized = FALSE
  )
  
  # Detect file type
  output$file_type <- reactive({
    req(input$file)
    ext <- tools::file_ext(input$file$name)
    ifelse(ext == "csv", "csv", "other")
  })
  outputOptions(output, "file_type", suspendWhenHidden = FALSE)
  
  # Load data based on detected file type
  observeEvent(input$load_data, {
    req(input$file)
    
    tryCatch({
      withProgress(message = 'Loading data...', value = 0.5, {
        ext <- tools::file_ext(input$file$name)
        
        if (ext == "csv") {
          rv$data <- read.csv(
            input$file$datapath,
            header = input$header,
            sep = input$sep,
            quote = input$quote,
            stringsAsFactors = TRUE
          )
        } else if (ext %in% c("xls", "xlsx")) {
          rv$data <- read_excel(input$file$datapath)
        } else if (ext == "dta") {
          rv$data <- read_dta(input$file$datapath)
        } else if (ext == "sav") {
          rv$data <- read_sav(input$file$datapath)
        } else {
          stop("Unsupported file type")
        }
        
        # Convert character columns to factor
        char_cols <- sapply(rv$data, is.character)
        rv$data[char_cols] <- lapply(rv$data[char_cols], as.factor)
        
        incProgress(1, detail = "Data loaded successfully!")
        Sys.sleep(0.5)
      })
      
      showNotification("Data loaded successfully!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # Data preview
  output$data_preview <- renderDT({
    req(rv$data)
    datatable(
      rv$data,
      options = list(
        scrollX = TRUE,
        pageLength = 5,
        dom = 'tip',
        initComplete = JS(
          "function(settings, json) {",
          "$(this.api().table().header()).css({'background-color': '#2c7a4d', 'color': '#fff'});",
          "}"
        )
      ),
      class = 'hover row-border',
      rownames = FALSE
    ) %>% 
      formatStyle(names(rv$data), backgroundColor = 'white')
  })
  
  # Dynamic UI for variable selection
  output$time_var_selector <- renderUI({
    req(rv$data)
    selectInput("time_var", "Time variable", choices = names(rv$data))
  })
  
  output$event_var_selector <- renderUI({
    req(rv$data)
    selectInput("event_var", "Event variable", choices = names(rv$data))
  })
  
  output$group_var_selector <- renderUI({
    req(rv$data)
    selectInput("group_var", "Grouping variable (optional)", 
                choices = c("None", names(rv$data)[sapply(rv$data, function(x) length(unique(x)) < 10)]))
  })
  
  # Color pickers for groups
  output$color_pickers <- renderUI({
    req(input$group_var != "None", rv$group_levels)
    
    lapply(seq_along(rv$group_levels), function(i) {
      colourInput(
        inputId = paste0("color_", i),
        label = paste("Color for", rv$group_levels[i]),
        value = if (i <= 8) {
          paletteer::paletteer_d("ggsci::category10_d3")[i]
        } else {
          paletteer::paletteer_d("viridis::viridis", n = length(rv$group_levels))[i]
        }
      )
    })
  })
  
  # Get custom colors
  custom_colors <- reactive({
    req(input$group_var != "None", rv$group_levels)
    
    colors <- sapply(seq_along(rv$group_levels), function(i) {
      input[[paste0("color_", i)]]
    })
    
    if (any(sapply(colors, is.null))) {
      return(NULL)
    }
    
    colors
  })
  
  # Generate survival plot
  generate_surv_plot <- function() {
    req(rv$fit, rv$data)
    
    # Create a list of plot parameters
    params <- list(
      fit = rv$fit,
      data = rv$data,
      conf.int = input$conf_int,
      pval = input$show_pval,
      risk.table = input$risk_table,
      title = input$plot_title,
      xlab = input$xlab,
      ylab = input$ylab,
      legend.title = input$legend_title,
      surv.median.line = ifelse(input$median_line, "hv", "none"),
      cumcens = input$cumulative_events,
      censor = input$censor_marks,
      ggtheme = switch(input$theme,
                       "minimal" = theme_minimal(base_size = input$font_size),
                       "classic" = theme_classic(base_size = input$font_size),
                       "gray" = theme_gray(base_size = input$font_size),
                       "bw" = theme_bw(base_size = input$font_size),
                       "light" = theme_light(base_size = input$font_size),
                       "dark" = theme_dark(base_size = input$font_size)),
      font.main = c(input$font_size, "bold"),
      font.x = c(input$font_size, "plain"),
      font.y = c(input$font_size, "plain"),
      font.tickslab = c(input$font_size, "plain"),
      font.legend = c(input$font_size, "plain")
    )
    
    # Add palette if grouping variable is specified and colors are available
    if (input$group_var != "None" && !is.null(custom_colors())) {
      params$palette <- custom_colors()
    }
    
    # Generate the plot
    p <- do.call(ggsurvplot, params)
    
    # Return the plot object
    p
  }
  
  # Observe changes that should trigger plot regeneration
  observe({
    # List of inputs that should trigger plot regeneration
    input$generate_plot
    input$conf_int
    input$show_pval
    input$risk_table
    input$plot_title
    input$xlab
    input$ylab
    input$legend_title
    input$font_size
    input$theme
    input$median_line
    input$cumulative_events
    input$censor_marks
    custom_colors()
    
    # Only regenerate if plot has been initialized
    if (rv$plot_initialized) {
      output$surv_plot <- renderGirafe({
        p <- generate_surv_plot()
        ggiraph::girafe(ggobj = p$plot, width_svg = 10, height_svg = 6)
      })
    }
  })
  
  # Main plot generation event
  observeEvent(input$generate_plot, {
    req(rv$data, input$time_var, input$event_var)
    
    tryCatch({
      # Data preparation
      rv$data[[input$time_var]] <- as.numeric(rv$data[[input$time_var]])
      rv$data[[input$event_var]] <- as.numeric(as.character(rv$data[[input$event_var]]))
      rv$data <- rv$data[!is.na(rv$data[[input$time_var]]) & !is.na(rv$data[[input$event_var]]), ]
      
      if (input$group_var != "None") {
        rv$data[[input$group_var]] <- as.factor(rv$data[[input$group_var]])
        rv$group_levels <- levels(rv$data[[input$group_var]])
      }
      
      # Create survival formula
      surv_formula <- as.formula(paste("Surv(", input$time_var, ",", input$event_var, ") ~", 
                                       ifelse(input$group_var == "None", "1", input$group_var)))
      
      # Fit survival model
      fit <- surv_fit(surv_formula, data = rv$data)
      rv$fit <- fit
      
      # Generate initial plot
      output$surv_plot <- renderGirafe({
        p <- generate_surv_plot()
        ggiraph::girafe(ggobj = p$plot, width_svg = 10, height_svg = 6)
      })
      
      # Set flag that plot has been initialized
      rv$plot_initialized <- TRUE
      
      # Calculate log-rank test if grouping variable is specified
      if (input$group_var != "None") {
        rv$logrank <- survdiff(surv_formula, data = rv$data)
      }
      
    }, error = function(e) {
      showNotification(paste("Error generating plot:", e$message), type = "error")
    })
  })
  
  # Display log-rank test results
  output$logrank_test <- renderPrint({
    req(rv$logrank)
    cat("Log-Rank Test for Equality of Survival Curves\n\n")
    print(rv$logrank)
    cat("\n\np-value:", 1 - pchisq(rv$logrank$chisq, length(rv$logrank$n) - 1))
  })
  
  # Download log-rank test results
  output$download_logrank <- downloadHandler(
    filename = function() {
      paste("logrank_test_results_", Sys.Date(), ".txt", sep = "")
    },
    content = function(file) {
      req(rv$logrank)
      results <- capture.output({
        cat("Log-Rank Test for Equality of Survival Curves\n\n")
        print(rv$logrank)
        cat("\n\np-value:", 1 - pchisq(rv$logrank$chisq, length(rv$logrank$n) - 1))
      })
      writeLines(results, file)
    }
  )
  
  # Generate survival table
  observeEvent(input$generate_table, {
    req(rv$fit, input$time_points)
    
    tryCatch({
      withProgress(message = 'Generating survival table...', value = 0.5, {
        time_points <- as.numeric(unlist(strsplit(input$time_points, ",\\s*")))
        
        sumry <- summary(rv$fit, times = time_points, extend = TRUE)
        
        if (is.null(rv$fit$strata)) {
          surv_table <- data.frame(
            Time = sumry$time,
            Survival = round(sumry$surv, 3),
            Events = sumry$n.event,
            AtRisk = sumry$n.risk
          )
          
          if (input$show_ci_table) {
            surv_table$LowerCI <- round(sumry$lower, 3)
            surv_table$UpperCI <- round(sumry$upper, 3)
          }
        } else {
          surv_table <- data.frame(
            Group = rep(names(rv$fit$strata), each = length(sumry$time)/length(rv$fit$strata)),
            Time = sumry$time,
            Survival = round(sumry$surv, 3),
            Events = sumry$n.event,
            AtRisk = sumry$n.risk
          )
          
          if (input$show_ci_table) {
            surv_table$LowerCI <- round(sumry$lower, 3)
            surv_table$UpperCI <- round(sumry$upper, 3)
          }
        }
        
        if (input$show_median) {
          median_surv <- surv_median(rv$fit)
          if (!is.null(rv$fit$strata)) {
            median_surv$Group <- gsub(".*=", "", median_surv$strata)
          }
          surv_table <- merge(surv_table, median_surv, all.x = TRUE)
        }
        
        rv$surv_table <- surv_table
        
        incProgress(1, detail = "Table generated!")
      })
    }, error = function(e) {
      showNotification(paste("Error generating table:", e$message), type = "error")
    })
  })
  
  # Display survival table
  output$surv_table <- renderDT({
    req(rv$surv_table)
    datatable(
      rv$surv_table,
      options = list(
        scrollX = TRUE,
        pageLength = 10,
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel'),
        initComplete = JS(
          "function(settings, json) {",
          "$(this.api().table().header()).css({'background-color': '#2c7a4d', 'color': '#fff'});",
          "}"
        )
      ),
      class = 'hover row-border',
      rownames = FALSE,
      extensions = 'Buttons'
    ) %>% 
      formatStyle(names(rv$surv_table), backgroundColor = 'white')
  })
  
  # Download survival table as Word document
  output$download_table_docx <- downloadHandler(
    filename = function() {
      paste("survival_table_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(rv$surv_table)
      
      ft <- flextable(rv$surv_table) %>%
        theme_box() %>%
        autofit() %>%
        bg(part = "header", bg = "#2c7a4d") %>%
        color(part = "header", color = "white") %>%
        bold(part = "header") %>%
        align(part = "all", align = "center")
      
      doc <- read_docx() %>%
        body_add_flextable(ft)
      
      print(doc, target = file)
    }
  )
  
  # Download plot as PNG
  output$download_plot_png <- downloadHandler(
    filename = function() {
      paste("survival_plot_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(rv$fit)
      
      p <- generate_surv_plot()
      
      ggsave(
        file,
        plot = print(p),
        device = "png",
        width = 12,
        height = 8,
        dpi = 600,
        bg = "white"
      )
    }
  )
  
  # Download plot as PDF
  output$download_plot_pdf <- downloadHandler(
    filename = function() {
      paste("survival_plot_", Sys.Date(), ".pdf", sep = "")
    },
    content = function(file) {
      req(rv$fit)
      
      p <- generate_surv_plot()
      
      ggsave(
        file,
        plot = print(p),
        device = "pdf",
        width = 12,
        height = 8,
        bg = "white"
      )
    }
  )
  
  # Download plot as SVG
  output$download_plot_svg <- downloadHandler(
    filename = function() {
      paste("survival_plot_", Sys.Date(), ".svg", sep = "")
    },
    content = function(file) {
      req(rv$fit)
      
      p <- generate_surv_plot()
      
      ggsave(
        file,
        plot = print(p),
        device = "svg",
        width = 12,
        height = 8,
        bg = "white"
      )
    }
  )
}

shinyApp(ui, server)
