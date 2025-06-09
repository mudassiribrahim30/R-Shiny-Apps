library(shiny)
library(shinythemes)
library(DT)
library(ggplot2)
library(dplyr)
library(stats)
library(haven)
library(readxl)
library(officer)
library(flextable)
library(ggiraph)
library(shinyjs)
library(shinyWidgets)
library(patchwork)

ui <- fluidPage(
  theme = shinytheme("flatly"),
  tags$head(tags$style(HTML("
    :root {
      --primary: #4a6fa5;
      --secondary: #166088;
      --accent: #4fc3f7;
      --dark: #2d3748;
      --light: #f8f9fa;
      --success: #4db6ac;
      --warning: #ff7043;
    }
    
    body {
      background-color: #f8f9fa;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    .navbar {
      background: linear-gradient(135deg, var(--primary), var(--secondary));
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    
    .navbar-brand {
      color: white !important;
      font-weight: 600;
    }
    
    .well {
      background-color: white;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      border: none;
      margin-bottom: 15px;
    }
    
    .btn-primary {
      background-color: var(--primary);
      border-color: var(--primary);
      transition: all 0.3s ease;
    }
    
    .btn-primary:hover {
      background-color: var(--secondary);
      border-color: var(--secondary);
      transform: translateY(-1px);
      box-shadow: 0 4px 8px rgba(22, 96, 136, 0.3);
    }
    
    .footer {
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      background: linear-gradient(90deg, var(--dark), var(--secondary));
      color: white;
      padding: 10px 20px;
      text-align: center;
      font-size: 0.9em;
      z-index: 1000;
      box-shadow: 0 -2px 10px rgba(0,0,0,0.1);
    }
    
    .footer a {
      color: var(--accent) !important;
      text-decoration: none;
      font-weight: 500;
    }
    
    .footer a:hover {
      color: var(--success) !important;
      text-decoration: underline;
    }
    
    .tab-content {
      background-color: white;
      border-radius: 8px;
      padding: 20px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      margin-bottom: 60px; /* Space for footer */
    }
    
    .slider-handle {
      background: var(--primary);
    }
    
    .irs-bar {
      background: linear-gradient(90deg, var(--accent), var(--success));
    }
  "))),
  
  # Header with gradient
  navbarPage(
    title = div(
      span("RegEffectXplorer", style = "color: white; font-weight: 600;"),
      span("", style = "color: rgba(255,255,255,0.7); font-size: 0.7em; margin-left: 5px;")
    ),
    windowTitle = "Regression Explorer",
    id = "nav",
    collapsible = TRUE,
    fluid = TRUE
  ),
  
  # Main content
  sidebarLayout(
    sidebarPanel(
      width = 4,
      style = "height: calc(100vh - 120px); overflow-y: auto; padding-right: 10px;",
      
      # Data Input Section
      div(
        class = "well",
        h4(icon("database"), " Data Input", style = "color: var(--dark);"),
        radioGroupButtons(
          inputId = "data_source",
          label = "Data Source",
          choices = c("CSV", "Excel", "SPSS", "Stata"),
          selected = "CSV",
          status = "primary"
        ),
        conditionalPanel(
          condition = "input.data_source == 'CSV'",
          fileInput("file_csv", "Upload CSV File", accept = ".csv"),
          checkboxInput("header", "Header", TRUE)
        ),
        conditionalPanel(
          condition = "input.data_source == 'Excel'",
          fileInput("file_excel", "Upload Excel File", accept = c(".xls", ".xlsx")),
          numericInput("sheet_num", "Sheet Number", value = 1)
        ),
        conditionalPanel(
          condition = "input.data_source == 'SPSS'",
          fileInput("file_spss", "Upload SPSS File", accept = ".sav")
        ),
        conditionalPanel(
          condition = "input.data_source == 'Stata'",
          fileInput("file_stata", "Upload Stata File", accept = ".dta")
        )
      ),
      
      # Variable Selection
      div(
        class = "well",
        h4(icon("sliders-h"), " Variable Selection", style = "color: var(--dark);"),
        selectInput("dep_var", "Dependent Variable", choices = NULL),
        selectizeInput("indep_vars", "Independent Variables", choices = NULL, multiple = TRUE),
        radioButtons("model_type", "Model Type",
                     choices = c("Linear Regression" = "linear", 
                                 "Poisson Regression" = "poisson"),
                     selected = "linear"),
        uiOutput("cat_var_ui")
      ),
      
      # Model Controls
      div(
        class = "well",
        h4(icon("calculator"), " Model Controls", style = "color: var(--dark);"),
        actionButton("run_model", "Run Analysis", 
                     icon = icon("rocket"),
                     class = "btn-primary btn-block",
                     style = "font-weight: 600;"),
        uiOutput("var_sliders"),
        uiOutput("cat_var_sliders")
      ),
      
      # Export Section
      div(
        class = "well",
        h4(icon("download"), " Export Results", style = "color: var(--dark);"),
        downloadButton("download_report", "Download Report", 
                       class = "btn-block",
                       style = "background-color: var(--dark); border-color: var(--dark);"),
        br(),
        downloadButton("download_plot", "Download Plots", 
                       class = "btn-block",
                       style = "background-color: var(--accent); border-color: var(--accent);")
      )
    ),
    
    # Main panel
    mainPanel(
      width = 8,
      style = "height: calc(100vh - 120px); overflow-y: auto;",
      tabsetPanel(
        id = "main_tabs",
        tabPanel(icon("table"), "Data Preview", DTOutput("data_preview")),
        tabPanel(icon("chart-line"), "Visualization", 
                 h3("Analysis Results", style = "color: var(--dark);"),
                 uiOutput("univariate_plots"),
                 h4("Multivariate Exploration", style = "margin-top: 20px;"),
                 girafeOutput("multivariate_plot"),
                 verbatimTextOutput("multivariate_prediction")),
        tabPanel(icon("list-alt"), "Model Summary", 
                 verbatimTextOutput("model_summary")),
        tabPanel(icon("table"), "Coefficients", 
                 DTOutput("coef_table"))
      )
    )
  ),
  
  # Stylish footer
  div(class = "footer",
      HTML("Developed with <span style='color:var(--warning);'>♥</span> by "),
      tags$a(href = "mailto:mudassiribrahim30@gmail.com", "Mudasir Mohammed Ibrahim"),
      HTML(" &bull; "),
      tags$a(href = "mailto:mudassiribrahim30@gmail.com?subject=Regression%20Explorer%20Feedback", 
             "Contact for support/issues")
  )
)

server <- function(input, output, session) {
  
  # Reactive data with error handling
  data <- reactive({
    req(input$data_source)
    
    tryCatch({
      if (input$data_source == "CSV" && !is.null(input$file_csv)) {
        read.csv(input$file_csv$datapath, header = input$header)
      } else if (input$data_source == "Excel" && !is.null(input$file_excel)) {
        read_excel(input$file_excel$datapath, sheet = input$sheet_num)
      } else if (input$data_source == "SPSS" && !is.null(input$file_spss)) {
        read_sav(input$file_spss$datapath)
      } else if (input$data_source == "Stata" && !is.null(input$file_stata)) {
        read_dta(input$file_stata$datapath)
      } else {
        NULL
      }
    }, error = function(e) {
      showNotification(paste("Error loading data:", e$message), type = "error")
      NULL
    })
  })
  
  # Update variable selections with validation
  observe({
    df <- data()
    req(df)
    
    # Ensure df is data.frame and has columns
    if (!is.data.frame(df)) {
      showNotification("Invalid data format", type = "error")
      return()
    }
    
    if (ncol(df) < 1) {
      showNotification("No columns found in data", type = "error")
      return()
    }
    
    var_choices <- names(df)
    updateSelectInput(session, "dep_var", choices = var_choices)
    updateSelectizeInput(session, "indep_vars", choices = var_choices)
  })
  
  # Safe categorical variable identification
  cat_vars <- reactive({
    req(input$indep_vars, data())
    df <- data()
    
    # Handle case where selected vars don't exist
    selected_vars <- intersect(input$indep_vars, names(df))
    if (length(selected_vars) == 0) return(NULL)
    
    cat_vars <- selected_vars[sapply(df[, selected_vars, drop = FALSE], function(x) {
      is.character(x) || is.factor(x) || (length(unique(na.omit(x))) < 5)
    })]
    
    if (length(cat_vars) == 0) return(NULL)
    cat_vars
  })
  
  
  # Robust UI for reference categories
  output$cat_var_ui <- renderUI({
    vars <- cat_vars()
    req(vars, data())
    df <- data()
    
    ref_ui <- tagList(h5("Reference Categories:"))
    
    for (var in vars) {
      if (!var %in% names(df)) next
      
      levels <- tryCatch({
        if (is.factor(df[[var]])) {
          levels(df[[var]])
        } else if (is.character(df[[var]])) {
          unique(na.omit(df[[var]]))
        } else if (length(unique(na.omit(df[[var]]))) < 5) {
          unique(na.omit(df[[var]]))
        } else {
          NULL
        }
      }, error = function(e) NULL)
      
      if (is.null(levels)) next
      
      ref_ui <- tagAppendChild(
        ref_ui,
        selectInput(
          paste0("ref_", var), 
          label = paste("Reference for", var), 
          choices = levels,
          selected = levels[1]
        )
      )
    }
    
    ref_ui
  })
  
  # Safe model data preparation
  model_data <- reactive({
    req(input$dep_var, input$indep_vars, data())
    df <- data()
    
    # Check if selected vars exist
    if (!all(c(input$dep_var, input$indep_vars) %in% names(df))) {
      showNotification("Selected variables not found in data", type = "error")
      return(NULL)
    }
    
    tryCatch({
      model_df <- df[, c(input$dep_var, input$indep_vars), drop = FALSE]
      
      # Process categorical variables if they exist
      vars <- cat_vars()
      if (!is.null(vars)) {
        for (var in vars) {
          if (!var %in% names(model_df)) next
          ref <- input[[paste0("ref_", var)]]
          if (is.null(ref)) next
          
          # Convert to factor if not already
          if (!is.factor(model_df[[var]])) {
            model_df[[var]] <- as.factor(model_df[[var]])
          }
          
          # Relevel if reference exists
          if (ref %in% levels(model_df[[var]])) {
            model_df[[var]] <- relevel(model_df[[var]], ref = ref)
          }
        }
      }
      
      model_df
    }, error = function(e) {
      showNotification(paste("Error preparing data:", e$message), type = "error")
      NULL
    })
  })
  
  # Robust model fitting
  model <- eventReactive(input$run_model, {
    req(model_data(), input$dep_var, input$indep_vars)
    df <- model_data()
    
    # Check for missing values in dependent variable
    if (any(is.na(df[[input$dep_var]]))) {
      showNotification("Dependent variable contains missing values", type = "warning")
    }
    
    tryCatch({
      formula <- as.formula(paste(input$dep_var, "~", paste(input$indep_vars, collapse = " + ")))
      
      if (input$model_type == "linear") {
        lm(formula, data = df)
      } else if (input$model_type == "poisson") {
        # Check if dependent variable is non-negative for Poisson
        if (any(df[[input$dep_var]] < 0)) {
          showNotification("Poisson regression requires non-negative dependent variable", 
                           type = "error")
          return(NULL)
        }
        glm(formula, data = df, family = poisson(link = "log"))
      }
    }, error = function(e) {
      showNotification(paste("Error fitting model:", e$message), type = "error")
      NULL
    })
  })
  
  # Safe model summary
  output$model_summary <- renderPrint({
    mod <- model()
    req(mod)
    tryCatch({
      summary(mod)
    }, error = function(e) {
      cat("Error generating summary:", e$message)
    })
  })
  
  # Safe coefficients table
  output$coef_table <- renderDT({
    mod <- model()
    req(mod)
    
    tryCatch({
      coef_df <- as.data.frame(summary(mod)$coefficients)
      coef_df <- round(coef_df, 4)
      coef_df$Significance <- ifelse(coef_df[,4] < 0.001, "***",
                                     ifelse(coef_df[,4] < 0.01, "**",
                                            ifelse(coef_df[,4] < 0.05, "*",
                                                   ifelse(coef_df[,4] < 0.1, ".", ""))))
      datatable(coef_df, options = list(pageLength = 10, dom = 'tip')) %>%
        formatStyle(names(coef_df), backgroundColor = 'white')
    }, error = function(e) {
      datatable(data.frame(Error = "Could not generate coefficients"))
    })
  })
  
  # Safe slider creation for numeric variables
  output$var_sliders <- renderUI({
    req(model(), input$indep_vars, model_data())
    df <- model_data()
    
    num_vars <- input$indep_vars[sapply(df[, input$indep_vars, drop = FALSE], is.numeric)]
    if (length(num_vars) == 0) return(NULL)
    
    slider_ui <- tagList(h4("Adjust Numeric Variables:"))
    
    for (var in num_vars) {
      var_values <- df[[var]]
      if (all(is.na(var_values))) next
      
      slider_ui <- tagAppendChild(
        slider_ui,
        sliderInput(
          inputId = paste0("slider_", var),
          label = var,
          min = floor(min(var_values, na.rm = TRUE)),
          max = ceiling(max(var_values, na.rm = TRUE)),
          value = median(var_values, na.rm = TRUE),
          step = ifelse(diff(range(var_values, na.rm = TRUE)) > 10, 1, 0.1)
        )
      )
    }
    
    slider_ui
  })
  
  # Safe selector creation for categorical variables
  output$cat_var_sliders <- renderUI({
    vars <- cat_vars()
    req(vars, model_data())
    df <- model_data()
    
    if (length(vars) == 0) return(NULL)
    
    slider_ui <- tagList(h4("Adjust Categorical Variables:"))
    
    for (var in vars) {
      if (!var %in% names(df)) next
      
      levels <- tryCatch({
        if (is.factor(df[[var]])) levels(df[[var]])
        else unique(na.omit(df[[var]]))
      }, error = function(e) NULL)
      
      if (is.null(levels)) next
      
      slider_ui <- tagAppendChild(
        slider_ui,
        selectInput(
          inputId = paste0("cat_slider_", var),
          label = var,
          choices = levels,
          selected = levels[1]
        )
      )
    }
    
    slider_ui
  })
  
  # Safe univariate plots generation
  output$univariate_plots <- renderUI({
    req(model(), input$indep_vars, model_data())
    df <- model_data()
    vars <- input$indep_vars
    
    plot_output_list <- lapply(seq_along(vars), function(i) {
      var <- vars[i]
      plotname <- paste0("plot_", var)
      
      div(
        class = "plot-container",
        div(class = "plot-title", paste("Effect of", var, "on", input$dep_var)),
        girafeOutput(plotname, width = "100%", height = "300px"),
        div(
          class = "prediction-box",
          verbatimTextOutput(paste0("prediction_", var))
        )
      )
    })
    
    do.call(tagList, plot_output_list)
  })
  
  # Render individual univariate plots with error handling
  observe({
    req(model(), input$indep_vars, model_data())
    df <- model_data()
    vars <- input$indep_vars
    
    for (var in vars) {
      local({
        local_var <- var
        
        tryCatch({
          if (is.numeric(df[[local_var]])) {
            # Numeric variable handling - create prediction data
            plot_data <- data.frame(x = seq(min(df[[local_var]], na.rm = TRUE), 
                                            max(df[[local_var]], na.rm = TRUE), 
                                            length.out = 100))
            names(plot_data) <- local_var
            
            # Set other variables to their reference values
            for (v in setdiff(vars, local_var)) {
              if (is.numeric(df[[v]])) {
                plot_data[[v]] <- median(df[[v]], na.rm = TRUE)
              } else {
                plot_data[[v]] <- levels(df[[v]])[1]
              }
            }
            
            # Generate predictions
            if (input$model_type == "linear") {
              pred <- predict(model(), newdata = plot_data, interval = "confidence")
              plot_data$pred <- pred[, "fit"]
              plot_data$lower <- pred[, "lwr"]
              plot_data$upper <- pred[, "upr"]
            } else {
              # Poisson regression - predict on response scale
              pred <- predict(model(), newdata = plot_data, type = "response", se.fit = TRUE)
              plot_data$pred <- pred$fit
              plot_data$lower <- pred$fit - 1.96 * pred$se.fit
              plot_data$upper <- pred$fit + 1.96 * pred$se.fit
            }
            
            # Get current slider value or use median if slider doesn't exist yet
            current_x <- if (paste0("slider_", local_var) %in% names(input)) {
              input[[paste0("slider_", local_var)]]
            } else {
              median(df[[local_var]], na.rm = TRUE)
            }
            
            # Create prediction data for current value
            current_data <- plot_data[which.min(abs(plot_data[[local_var]] - current_x)), ]
            
            # Create the plot
            p <- ggplot(plot_data, aes_string(x = local_var, y = "pred")) +
              geom_ribbon(aes(ymin = lower, ymax = upper), alpha = 0.2, fill = "#3498db") +
              geom_line(size = 1.5, color = "#2c3e50") +
              geom_vline(xintercept = current_x, linetype = "dashed", color = "#e74c3c") +
              geom_point_interactive(aes(x = current_x, y = current_data$pred, 
                                         tooltip = paste0("Value: ", round(current_x, 2), "\n",
                                                          input$dep_var, ": ", round(current_data$pred, 2)),
                                         data_id = paste0("drag_", local_var)),
                                     size = 3, color = "#e74c3c") +
              labs(y = input$dep_var, x = local_var) +
              theme_minimal(base_size = 12) +
              theme(panel.grid.major = element_line(color = "#f0f0f0"),
                    panel.grid.minor = element_blank(),
                    panel.background = element_rect(fill = "white", color = NA))
            
            output[[paste0("plot_", local_var)]] <- renderGirafe({
              girafe(ggobj = p, width_svg = 8, height_svg = 3,
                     options = list(
                       opts_hover(css = "fill:#2ecc71;stroke:#27ae60;"),
                       opts_selection(type = "none"),
                       opts_tooltip(css = "background-color:white;color:black;padding:5px;border-radius:3px;box-shadow: 0 0 5px rgba(0,0,0,0.2);"),
                       opts_sizing(rescale = TRUE)
                     ))
            })
            
            # Handle dragging points in univariate plots
            observeEvent(input[[paste0("plot_", local_var, "_selected")]], {
              selected_point <- input[[paste0("plot_", local_var, "_selected")]]
              if (!is.null(selected_point) && selected_point == paste0("drag_", local_var)) {
                new_x <- input[[paste0("plot_", local_var, "_hover")]]$x
                if (!is.null(new_x)) {
                  updateSliderInput(session, paste0("slider_", local_var), value = new_x)
                }
              }
            })
            
            # Current prediction output
            output[[paste0("prediction_", local_var)]] <- renderPrint({
              new_data <- df[1, , drop = FALSE]
              new_data[1, ] <- NA
              
              # Set values based on sliders or defaults
              for (v in vars) {
                if (v == local_var) {
                  new_data[[v]] <- current_x
                } else {
                  if (is.numeric(df[[v]])) {
                    new_data[[v]] <- median(df[[v]], na.rm = TRUE)
                  } else {
                    new_data[[v]] <- levels(df[[v]])[1]
                  }
                }
              }
              
              if (input$model_type == "linear") {
                pred <- predict(model(), newdata = new_data, interval = "confidence")
                cat("Current", local_var, "value:", round(current_x, 2), "\n")
                cat("Predicted", input$dep_var, ":", round(pred[1], 3), "\n")
                cat("95% CI: [", round(pred[2], 3), ", ", round(pred[3], 3), "]", sep = "")
              } else {
                pred <- predict(model(), newdata = new_data, type = "response", se.fit = TRUE)
                ci_lower <- pred$fit - 1.96 * pred$se.fit
                ci_upper <- pred$fit + 1.96 * pred$se.fit
                cat("Current", local_var, "value:", round(current_x, 2), "\n")
                cat("Predicted", input$dep_var, ":", round(pred$fit, 3), "\n")
                cat("95% CI: [", round(ci_lower, 3), ", ", round(ci_upper, 3), "]", sep = "")
              }
            })
            
          } else {
            # Categorical variable handling
            levels <- levels(df[[local_var]])
            
            plot_data <- do.call(rbind, lapply(levels, function(lvl) {
              temp <- df[1, , drop = FALSE]
              temp[1, ] <- NA
              temp[[local_var]] <- lvl
              
              for (v in setdiff(vars, local_var)) {
                if (is.numeric(df[[v]])) {
                  temp[[v]] <- median(df[[v]], na.rm = TRUE)
                } else {
                  temp[[v]] <- levels(df[[v]])[1]
                }
              }
              
              if (input$model_type == "linear") {
                pred <- predict(model(), newdata = temp, interval = "confidence")
                data.frame(level = lvl, 
                           pred = pred[1], 
                           lower = pred[2], 
                           upper = pred[3])
              } else {
                pred <- predict(model(), newdata = temp, type = "response", se.fit = TRUE)
                data.frame(level = lvl,
                           pred = pred$fit,
                           lower = pred$fit - 1.96 * pred$se.fit,
                           upper = pred$fit + 1.96 * pred$se.fit)
              }
            }))
            
            current_level <- if (paste0("cat_slider_", local_var) %in% names(input)) {
              input[[paste0("cat_slider_", local_var)]]
            } else {
              levels[1]
            }
            
            # Create interactive plot
            p <- ggplot(plot_data, aes(x = level, y = pred)) +
              geom_col(fill = "#3498db", alpha = 0.7) +
              geom_errorbar(aes(ymin = lower, ymax = upper), width = 0.2, color = "#2c3e50") +
              geom_point_interactive(
                aes(
                  tooltip = paste0("Level: ", level, "\n", input$dep_var, ": ", round(pred, 2))
                ),
                size = 3,
                color = ifelse(plot_data$level == current_level, "#e74c3c", "#3498db")
              ) +
              labs(y = input$dep_var, x = local_var) +
              theme_minimal(base_size = 12) +
              theme(
                panel.grid.major = element_line(color = "#f0f0f0"),
                panel.grid.minor = element_blank(),
                panel.background = element_rect(fill = "white", color = NA),
                axis.text.x = element_text(angle = 45, hjust = 1)
              )
            
            # Render the interactive plot
            output[[paste0("plot_", local_var)]] <- renderGirafe({
              girafe(
                ggobj = p, width_svg = 8, height_svg = 3,
                options = list(
                  opts_hover(css = "fill:#2ecc71;stroke:#27ae60;"),
                  opts_selection(type = "none"),
                  opts_tooltip(css = "background-color:white;color:black;padding:5px;border-radius:3px;box-shadow: 0 0 5px rgba(0,0,0,0.2);"),
                  opts_sizing(rescale = TRUE)
                )
              )
            })
            
            # Current prediction output
            output[[paste0("prediction_", local_var)]] <- renderPrint({
              new_data <- df[1, , drop = FALSE]
              new_data[1, ] <- NA
              new_data[[local_var]] <- current_level
              
              for (v in setdiff(vars, local_var)) {
                if (is.numeric(df[[v]])) {
                  new_data[[v]] <- median(df[[v]], na.rm = TRUE)
                } else {
                  new_data[[v]] <- levels(df[[v]])[1]
                }
              }
              
              if (input$model_type == "linear") {
                pred <- predict(model(), newdata = new_data, interval = "confidence")
                cat("Current", local_var, "level:", current_level, "\n")
                cat("Predicted", input$dep_var, ":", round(pred[1], 3), "\n")
                cat("95% CI: [", round(pred[2], 3), ", ", round(pred[3], 3), "]", sep = "")
              } else {
                pred <- predict(model(), newdata = new_data, type = "response", se.fit = TRUE)
                ci_lower <- pred$fit - 1.96 * pred$se.fit
                ci_upper <- pred$fit + 1.96 * pred$se.fit
                cat("Current", local_var, "level:", current_level, "\n")
                cat("Predicted", input$dep_var, ":", round(pred$fit, 3), "\n")
                cat("95% CI: [", round(ci_lower, 3), ", ", round(ci_upper, 3), "]", sep = "")
              }
            })
          }
        }, error = function(e) {
          output[[paste0("plot_", local_var)]] <- renderGirafe({
            girafe(ggobj = ggplot() + 
                     annotate("text", x = 0.5, y = 0.5, label = "Error generating plot", size = 6) +
                     theme_void(),
                   width_svg = 8, height_svg = 3)
          })
          
          output[[paste0("prediction_", local_var)]] <- renderPrint({
            cat("Error generating prediction")
          })
        })
      })
    }
  })
  
  # Safe multivariate plot
  output$multivariate_plot <- renderGirafe({
    req(model(), input$indep_vars, model_data())
    df <- model_data()
    num_vars <- input$indep_vars[sapply(df[, input$indep_vars, drop = FALSE], is.numeric)]
    
    if (length(num_vars) == 0) return(NULL)
    
    tryCatch({
      var <- num_vars[1]
      other_vars <- setdiff(input$indep_vars, var)
      
      # Create prediction data
      plot_data <- data.frame(x = seq(min(df[[var]], na.rm = TRUE), 
                                      max(df[[var]], na.rm = TRUE), 
                                      length.out = 100))
      names(plot_data) <- var
      
      # Set other variables to their current values (from sliders or defaults)
      for (v in other_vars) {
        if (v %in% cat_vars()) {
          if (paste0("cat_slider_", v) %in% names(input)) {
            plot_data[[v]] <- input[[paste0("cat_slider_", v)]]
          } else {
            plot_data[[v]] <- levels(df[[v]])[1]
          }
        } else if (is.numeric(df[[v]])) {
          if (paste0("slider_", v) %in% names(input)) {
            plot_data[[v]] <- input[[paste0("slider_", v)]]
          } else {
            plot_data[[v]] <- median(df[[v]], na.rm = TRUE)
          }
        }
      }
      
      # Generate predictions
      if (input$model_type == "linear") {
        pred <- predict(model(), newdata = plot_data, interval = "confidence")
        plot_data$pred <- pred[, "fit"]
        plot_data$lower <- pred[, "lwr"]
        plot_data$upper <- pred[, "upr"]
      } else {
        pred <- predict(model(), newdata = plot_data, type = "response", se.fit = TRUE)
        plot_data$pred <- pred$fit
        plot_data$lower <- pred$fit - 1.96 * pred$se.fit
        plot_data$upper <- pred$fit + 1.96 * pred$se.fit
      }
      
      # Get current slider value or use median if slider doesn't exist yet
      current_x <- if (paste0("slider_", var) %in% names(input)) {
        input[[paste0("slider_", var)]]
      } else {
        median(df[[var]], na.rm = TRUE)
      }
      
      # Create prediction data for current value
      current_data <- plot_data[which.min(abs(plot_data[[var]] - current_x)), ]
      
      # Create the plot
      p <- ggplot(plot_data, aes_string(x = var, y = "pred")) +
        geom_ribbon(aes(ymin = lower, ymax = upper), alpha = 0.2, fill = "#3498db") +
        geom_line(size = 1.5, color = "#2c3e50") +
        geom_point_interactive(aes(
          x = current_x,
          y = current_data$pred,
          tooltip = paste0("Value: ", round(current_x, 2), "\n",
                           input$dep_var, ": ", round(current_data$pred, 2)),
          data_id = "drag_point"
        ), size = 3, color = "#e74c3c") +
        labs(
          title = "Multivariate Effect Exploration",
          subtitle = paste("Effect of", var, "on", input$dep_var, "with other variables held constant"),
          y = input$dep_var,
          x = var
        ) +
        theme_minimal(base_size = 14) +
        theme(
          plot.title = element_text(hjust = 0.5, face = "bold"),
          plot.subtitle = element_text(hjust = 0.5),
          panel.grid.major = element_line(color = "#f0f0f0"),
          panel.grid.minor = element_blank(),
          panel.background = element_rect(fill = "white", color = NA),
          plot.background = element_rect(fill = "white", color = NA)
        )
      
      girafe(ggobj = p, width_svg = 8, height_svg = 5,
             options = list(
               opts_hover(css = "fill:#2ecc71;stroke:#27ae60;"),
               opts_selection(type = "none"),
               opts_tooltip(css = "background-color:white;color:black;padding:5px;border-radius:3px;box-shadow: 0 0 5px rgba(0,0,0,0.2);"),
               opts_sizing(rescale = TRUE)
             ))
    }, error = function(e) {
      girafe(ggobj = ggplot() + 
               annotate("text", x = 0.5, y = 0.5, label = "Error generating plot", size = 6) +
               theme_void(),
             width_svg = 8, height_svg = 5)
    })
  })
  
  # Handle dragging in multivariate plot
  observeEvent(input$multivariate_plot_selected, {
    req(input$multivariate_plot_selected == "drag_point", 
        input$multivariate_plot_hovered,
        model(), input$indep_vars, model_data())
    
    df <- model_data()
    num_vars <- input$indep_vars[sapply(df[, input$indep_vars, drop = FALSE], is.numeric)]
    var <- num_vars[1]
    
    new_x <- input$multivariate_plot_hovered$x
    if (!is.null(new_x)) {
      updateSliderInput(session, paste0("slider_", var), value = new_x)
    }
  })
  
  # Safe multivariate prediction
  output$multivariate_prediction <- renderPrint({
    req(model(), input$indep_vars, model_data())
    df <- model_data()
    
    tryCatch({
      new_data <- df[1, , drop = FALSE]
      new_data[1, ] <- NA
      
      # Set values based on sliders or defaults
      for (var in input$indep_vars) {
        if (is.numeric(df[[var]])) {
          if (paste0("slider_", var) %in% names(input)) {
            new_data[[var]] <- input[[paste0("slider_", var)]]
          } else {
            new_data[[var]] <- median(df[[var]], na.rm = TRUE)
          }
        } else if (is.factor(df[[var]])) {
          if (paste0("cat_slider_", var) %in% names(input)) {
            new_data[[var]] <- input[[paste0("cat_slider_", var)]]
          } else {
            new_data[[var]] <- levels(df[[var]])[1]
          }
        }
      }
      
      if (input$model_type == "linear") {
        pred <- predict(model(), newdata = new_data, interval = "confidence")
        cat("Current Multivariate Prediction\n")
        cat("-----------------------------\n")
        
        for (var in input$indep_vars) {
          if (is.numeric(df[[var]])) {
            cat(var, ":", round(new_data[[var]], 2), "\n")
          } else {
            cat(var, ":", as.character(new_data[[var]]), "\n")
          }
        }
        
        cat("\nPredicted", input$dep_var, ":", round(pred[1], 3), "\n")
        cat("95% Confidence Interval: [", round(pred[2], 3), ", ", round(pred[3], 3), "]", sep = "")
      } else {
        pred <- predict(model(), newdata = new_data, type = "response", se.fit = TRUE)
        ci_lower <- pred$fit - 1.96 * pred$se.fit
        ci_upper <- pred$fit + 1.96 * pred$se.fit
        
        cat("Current Multivariate Prediction\n")
        cat("-----------------------------\n")
        
        for (var in input$indep_vars) {
          if (is.numeric(df[[var]])) {
            cat(var, ":", round(new_data[[var]], 2), "\n")
          } else {
            cat(var, ":", as.character(new_data[[var]]), "\n")
          }
        }
        
        cat("\nPredicted", input$dep_var, ":", round(pred$fit, 3), "\n")
        cat("95% Confidence Interval: [", round(ci_lower, 3), ", ", round(ci_upper, 3), "]", sep = "")
      }
    }, error = function(e) {
      cat("Error generating prediction:", e$message)
    })
  })
  
  # Safe data preview
  output$data_preview <- renderDT({
    df <- data()
    req(df)
    
    tryCatch({
      datatable(df, 
                options = list(
                  pageLength = 5, 
                  scrollX = TRUE,
                  dom = 'tip',
                  initComplete = JS(
                    "function(settings, json) {",
                    "$(this.api().table().header()).css({'background-color': '#2c3e50', 'color': '#fff'});",
                    "}")
                )) %>%
        formatStyle(names(df), backgroundColor = 'white')
    }, error = function(e) {
      datatable(data.frame(Error = "Could not display data"))
    })
  })
  
  # Safe plot download
  output$download_plot <- downloadHandler(
    filename = function() {
      paste("regression-plot-", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(model(), input$indep_vars, model_data())
      
      tryCatch({
        df <- model_data()
        vars <- input$indep_vars
        
        plots <- lapply(vars, function(var) {
          if (is.numeric(df[[var]])) {
            # Numeric variable plot
            plot_data <- data.frame(x = seq(min(df[[var]], na.rm = TRUE), 
                                            max(df[[var]], na.rm = TRUE), 
                                            length.out = 100))
            names(plot_data) <- var
            
            for (v in setdiff(vars, var)) {
              if (is.numeric(df[[v]])) {
                plot_data[[v]] <- median(df[[v]], na.rm = TRUE)
              } else {
                plot_data[[v]] <- levels(df[[v]])[1]
              }
            }
            
            if (input$model_type == "linear") {
              pred <- predict(model(), newdata = plot_data, interval = "confidence")
              plot_data$pred <- pred[, "fit"]
              plot_data$lower <- pred[, "lwr"]
              plot_data$upper <- pred[, "upr"]
            } else {
              pred <- predict(model(), newdata = plot_data, type = "response", se.fit = TRUE)
              plot_data$pred <- pred$fit
              plot_data$lower <- pred$fit - 1.96 * pred$se.fit
              plot_data$upper <- pred$fit + 1.96 * pred$se.fit
            }
            
            current_x <- if (paste0("slider_", var) %in% names(input)) {
              input[[paste0("slider_", var)]]
            } else {
              median(df[[var]], na.rm = TRUE)
            }
            
            current_data <- plot_data[which.min(abs(plot_data[[var]] - current_x)), ]
            
            ggplot(plot_data, aes_string(x = var, y = "pred")) +
              geom_ribbon(aes(ymin = lower, ymax = upper), alpha = 0.2, fill = "#3498db") +
              geom_line(size = 1.5, color = "#2c3e50") +
              geom_vline(xintercept = current_x, linetype = "dashed", color = "#e74c3c") +
              geom_point(aes(x = current_x, y = current_data$pred), size = 3, color = "#e74c3c") +
              labs(title = paste("Effect of", var, "on", input$dep_var),
                   y = input$dep_var,
                   x = var) +
              theme_minimal(base_size = 12) +
              theme(plot.title = element_text(hjust = 0.5, face = "bold"),
                    panel.grid.major = element_line(color = "#f0f0f0"),
                    panel.grid.minor = element_blank(),
                    panel.background = element_rect(fill = "white", color = NA))
          } else {
            # Categorical variable plot
            levels <- levels(df[[var]])
            
            plot_data <- do.call(rbind, lapply(levels, function(lvl) {
              temp <- df[1, , drop = FALSE]
              temp[1, ] <- NA
              temp[[var]] <- lvl
              
              for (v in setdiff(vars, var)) {
                if (is.numeric(df[[v]])) {
                  temp[[v]] <- median(df[[v]], na.rm = TRUE)
                } else {
                  temp[[v]] <- levels(df[[v]])[1]
                }
              }
              
              if (input$model_type == "linear") {
                pred <- predict(model(), newdata = temp, interval = "confidence")
                data.frame(
                  level = lvl, 
                  pred = pred[1], 
                  lower = pred[2], 
                  upper = pred[3]
                )
              } else {
                pred <- predict(model(), newdata = temp, type = "response", se.fit = TRUE)
                data.frame(
                  level = lvl,
                  pred = pred$fit,
                  lower = pred$fit - 1.96 * pred$se.fit,
                  upper = pred$fit + 1.96 * pred$se.fit
                )
              }
            }))
            
            current_level <- if (paste0("cat_slider_", var) %in% names(input)) {
              input[[paste0("cat_slider_", var)]]
            } else {
              levels[1]
            }
            
            ggplot(plot_data, aes(x = level, y = pred)) +
              geom_col(fill = "#3498db", alpha = 0.7) +
              geom_errorbar(aes(ymin = lower, ymax = upper), width = 0.2, color = "#2c3e50") +
              geom_point(aes(color = level == current_level), size = 3) +
              scale_color_manual(values = c("#3498db", "#e74c3c"), guide = FALSE) +
              labs(title = paste("Effect of", var, "on", input$dep_var),
                   y = input$dep_var,
                   x = var) +
              theme_minimal(base_size = 12) +
              theme(plot.title = element_text(hjust = 0.5, face = "bold"),
                    panel.grid.major = element_line(color = "#f0f0f0"),
                    panel.grid.minor = element_blank(),
                    panel.background = element_rect(fill = "white", color = NA),
                    axis.text.x = element_text(angle = 45, hjust = 1))
          }
        })
        
        # Combine plots
        combined_plot <- wrap_plots(plots, ncol = 1) +
          plot_annotation(title = "Regression Analysis Results",
                          theme = theme(plot.title = element_text(hjust = 0.5, face = "bold", size = 16)))
        
        ggsave(file, plot = combined_plot, device = "png", 
               width = 10, height = 3 * length(plots), dpi = 300, limitsize = FALSE)
      }, error = function(e) {
        showNotification(paste("Error saving plot:", e$message), type = "error")
      })
    }
  )
  
  # Safe Word report download
  output$download_report <- downloadHandler(
    filename = function() {
      paste("regression-report-", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(model())
      
      tryCatch({
        # Create a new docx file
        doc <- officer::read_docx()
        
        # Add title
        doc <- doc %>% 
          officer::body_add_par("Regression Analysis Report", style = "heading 1") %>%
          officer::body_add_par(paste("Generated on", format(Sys.Date(), "%B %d, %Y")), style = "Normal") %>%
          officer::body_add_par("", style = "Normal")
        
        # Add model specification
        doc <- doc %>%
          officer::body_add_par("Model Specification", style = "heading 2") %>%
          officer::body_add_par(paste("Model type:", ifelse(input$model_type == "linear", "Linear Regression", "Poisson Regression")), style = "Normal") %>%
          officer::body_add_par(paste("Dependent variable:", input$dep_var), style = "Normal") %>%
          officer::body_add_par(paste("Independent variables:", paste(input$indep_vars, collapse = ", ")), style = "Normal") %>%
          officer::body_add_par(paste("Model formula:", paste(input$dep_var, "~", paste(input$indep_vars, collapse = " + "))), style = "Normal") %>%
          officer::body_add_par("", style = "Normal")
        
        # Add model summary - fixed to handle multiple lines
        doc <- doc %>%
          officer::body_add_par("Model Summary", style = "heading 2")
        
        # Add summary line by line
        summary_lines <- capture.output(summary(model()))
        for (line in summary_lines) {
          doc <- doc %>% officer::body_add_par(line, style = "Normal")
        }
        doc <- doc %>% officer::body_add_par("", style = "Normal")
        
        # Add coefficients table
        coef_df <- as.data.frame(summary(model())$coefficients)
        coef_df <- round(coef_df, 4)
        coef_df$Significance <- ifelse(coef_df[,4] < 0.001, "***",
                                       ifelse(coef_df[,4] < 0.01, "**",
                                              ifelse(coef_df[,4] < 0.05, "*",
                                                     ifelse(coef_df[,4] < 0.1, ".", ""))))
        
        ft <- flextable::flextable(coef_df) %>%
          flextable::theme_box() %>%
          flextable::set_caption("Regression Coefficients") %>%
          flextable::autofit()
        
        doc <- doc %>%
          officer::body_add_par("Regression Coefficients", style = "heading 2") %>%
          flextable::body_add_flextable(ft) %>%
          officer::body_add_par("Signif. codes: 0 '***' 0.001 '**' 0.01 '*' 0.05 '.' 0.1 ' ' 1", style = "Normal") %>%
          officer::body_add_par("", style = "Normal")
        
        # Add interpretation
        doc <- doc %>%
          officer::body_add_par("Interpretation", style = "heading 2")
        
        # Interpret each coefficient
        for (i in 2:nrow(coef_df)) {
          var_name <- rownames(coef_df)[i]
          var <- gsub("^factor\\(|\\)$", "", var_name)
          coef <- coef_df[i, 1]
          pval <- coef_df[i, 4]
          
          interpretation <- paste0("• ", var_name, ": ")
          
          if (var %in% cat_vars()) {
            ref <- input[[paste0("ref_", var)]]
            interpretation <- paste0(interpretation, 
                                     "Compared to the reference category (", ref, "), ")
          } else {
            interpretation <- paste0(interpretation, "For each one-unit increase in ", var, ", ")
          }
          
          interpretation <- paste0(interpretation, 
                                   input$dep_var, " is expected to ")
          
          if (input$model_type == "linear") {
            if (coef > 0) {
              interpretation <- paste0(interpretation, "increase by ", abs(round(coef, 3)))
            } else {
              interpretation <- paste0(interpretation, "decrease by ", abs(round(coef, 3)))
            }
          } else {
            # For Poisson, coefficients are on log scale
            interpretation <- paste0(interpretation, "multiply by ", round(exp(coef), 3))
          }
          
          if (pval < 0.05) {
            interpretation <- paste0(interpretation, " (statistically significant, p = ", round(pval, 3), ")")
          } else {
            interpretation <- paste0(interpretation, " (not statistically significant, p = ", round(pval, 3), ")")
          }
          
          doc <- doc %>% officer::body_add_par(interpretation, style = "Normal")
        }
        
        # Add model fit interpretation
        if (input$model_type == "linear") {
          r_sq <- summary(model())$r.squared
          doc <- doc %>%
            officer::body_add_par("", style = "Normal") %>%
            officer::body_add_par(paste0("The model explains approximately ", round(r_sq * 100, 1), 
                                         "% of the variance in ", input$dep_var, "."), style = "Normal")
        } else {
          # For Poisson, we can add null and residual deviance
          deviance <- summary(model())$deviance
          null_deviance <- summary(model())$null.deviance
          pseudo_r2 <- 1 - (deviance / null_deviance)
          
          doc <- doc %>%
            officer::body_add_par("", style = "Normal") %>%
            officer::body_add_par(paste0("The model reduces deviance by approximately ", 
                                         round(pseudo_r2 * 100, 1), 
                                         "% compared to the null model."), style = "Normal")
        }
        
        # Save the document
        print(doc, target = file)
      }, error = function(e) {
        showNotification(paste("Error generating report:", e$message), type = "error")
      })
    }
  )
}

shinyApp(ui = ui, server = server)
