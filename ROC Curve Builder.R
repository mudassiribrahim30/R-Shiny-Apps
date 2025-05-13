# Load required libraries
library(shiny)
library(ggplot2)
library(plotly)
library(pROC)
library(MLmetrics)
library(DT)
library(caret)
library(dplyr)
library(readxl)
library(haven)
library(shinythemes)
library(officer)
library(flextable)

ui <- fluidPage(
  theme = shinytheme("flatly"),
  
  tags$head(
    tags$style(HTML("
      /* Main styling */
      body {
        background-color: #f5f7fa;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      }
      
      /* Header styling */
      .header {
        background: linear-gradient(135deg, #3498db, #2c3e50);
        color: white;
        padding: 20px 0;
        margin-bottom: 30px;
        border-radius: 0 0 10px 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
      }
      
      .app-title {
        font-size: 28px;
        font-weight: 700;
        margin: 0;
        text-shadow: 1px 1px 3px rgba(0,0,0,0.2);
        text-align: center;
      }
      
      .app-subtitle {
        font-size: 16px;
        opacity: 0.9;
        margin-top: 5px;
        text-align: center;
      }
      
      /* Panel styling */
      .sidebar-panel {
        background-color: white;
        border-radius: 10px;
        padding: 25px !important;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border: 1px solid #e0e0e0;
      }
      
      .main-panel {
        background-color: white;
        border-radius: 10px;
        padding: 25px !important;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border: 1px solid #e0e0e0;
      }
      
      /* Button styling */
      .btn-primary {
        background-color: #3498db;
        border-color: #3498db;
        font-weight: 500;
        border-radius: 5px;
        padding: 8px 15px;
        transition: all 0.3s;
      }
      
      .btn-primary:hover {
        background-color: #2980b9;
        border-color: #2980b9;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(41, 128, 185, 0.3);
      }
      
      .btn-success {
        background-color: #2ecc71;
        border-color: #2ecc71;
        font-weight: 500;
        border-radius: 5px;
        padding: 8px 15px;
        transition: all 0.3s;
      }
      
      .btn-success:hover {
        background-color: #27ae60;
        border-color: #27ae60;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(46, 204, 113, 0.3);
      }
      
      .btn-info {
        background-color: #17a2b8;
        border-color: #17a2b8;
        font-weight: 500;
        border-radius: 5px;
        padding: 8px 15px;
        transition: all 0.3s;
      }
      
      .btn-info:hover {
        background-color: #138496;
        border-color: #138496;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(23, 162, 184, 0.3);
      }
      
      /* Input styling */
      .shiny-input-container {
        margin-bottom: 15px;
      }
      
      .form-control {
        border-radius: 5px;
        border: 1px solid #ddd;
        padding: 8px 12px;
      }
      
      .selectize-input {
        border-radius: 5px !important;
        border: 1px solid #ddd !important;
        padding: 8px 12px !important;
        box-shadow: none !important;
      }
      
      /* Text styling */
      h2, h3, h4 {
        color: #2c3e50;
        font-weight: 600;
      }
      
      h4 {
        margin-top: 25px;
        border-bottom: 2px solid #f0f0f0;
        padding-bottom: 8px;
      }
      
      .input-info, .notice {
        font-size: 0.9em;
        color: #7f8c8d;
      }
      
      .instruction {
        font-size: 0.85em;
        color: #95a5a6;
        margin-bottom: 10px;
      }
      
      .warning-text {
        color: #e74c3c;
        font-weight: bold;
        background-color: #fde8e8;
        padding: 5px 10px;
        border-radius: 4px;
        display: inline-block;
        margin: 5px 0;
      }
      
      /* Footer styling */
      .footer {
        color: #95a5a6;
        font-style: italic;
        text-align: center;
        padding-top: 25px;
        margin-top: 20px;
        border-top: 1px solid #eee;
      }
      
      /* Table styling */
      .dataTables_wrapper {
        margin-top: 15px;
      }
      
      /* Responsive adjustments */
      @media (max-width: 768px) {
        .sidebar-panel, .main-panel {
          padding: 15px !important;
        }
      }
    "))
  ),
  
  # Header Section
  div(class = "header",
      div(class = "container-fluid",
          div(
            h1(class = "app-title", "ROC Curve Builder"),
            p(class = "app-subtitle", "Visualize and analyze binary classification performance")
          )
      )
  ),
  
  # Main Content
  div(class = "container-fluid",
      sidebarLayout(
        sidebarPanel(
          class = "sidebar-panel",
          width = 4,
          h3("Data Input", style = "margin-top: 0;"),
          
          fileInput("data", "Upload Your File",
                    accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta"),
                    buttonLabel = "Browse Files...",
                    placeholder = "No file selected",
                    width = "100%"),
          div(class = "input-info", "Supported formats: CSV, Excel (XLS/XLSX), Stata (DTA), SPSS (SAV)"),
          
          div(class = "notice", 
              icon("info-circle"), 
              strong(" Note:"), 
              " After uploading, click 'Generate ROC Curve' to populate variable selectors."),
          br(),
          
          h4("Variable Selection"),
          selectInput("outcome", "Outcome Variable", choices = NULL, width = "100%"),
          uiOutput("outcome_warning"),
          selectizeInput("predictors", "Predictor Variable(s)", choices = NULL, multiple = TRUE, width = "100%"),
          
          div(class = "instruction", 
              icon("exclamation-triangle"), 
              "Ensure outcome is binary (0/1). Recode if needed:"),
          uiOutput("recode_ui"),
          
          h4("Graph Customization"),
          textInput("graph_title", "Graph Title", "ROC Curve", width = "100%"),
          textInput("x_label", "X-axis Label", "False Positive Rate", width = "100%"),
          textInput("y_label", "Y-axis Label", "True Positive Rate", width = "100%"),
          
          h4("Actions"),
          div(style = "text-align: center;",
              actionButton("generate", "Generate ROC Curve", 
                           class = "btn-primary",
                           style = "width: 100%; margin-bottom: 10px;"),
              downloadButton("download_plot", "Download Plot (PNG)", 
                             class = "btn-success",
                             style = "width: 100%; margin-bottom: 10px;"),
              downloadButton("download_word", "Download Report (DOCX)", 
                             class = "btn-info",
                             style = "width: 100%;")
          )
        ),
        
        mainPanel(
          class = "main-panel",
          width = 8,
          plotlyOutput("roc_plot", height = "500px"),
          
          h4("Performance Metrics"),
          DTOutput("metrics_table"),
          
          h4("Confusion Matrix"),
          DTOutput("conf_matrix_table"),
          
          div(class = "footer",
              "ROC Curve Builder v1.1 | ",
              icon("envelope"), " ",
              a("mudassiribrahim30@gmail.com", href = "mailto:mudassiribrahim30@gmail.com"),
              " | ",
              icon("github"), " ",
              a("View on GitHub", href = "https://github.com", target = "_blank")
          )
        )
      )
  )
)

server <- function(input, output, session) {
  
  read_data <- reactive({
    req(input$data)
    ext <- tools::file_ext(input$data$datapath)
    
    df <- switch(tolower(ext),
                 csv = read.csv(input$data$datapath, stringsAsFactors = FALSE),
                 xlsx = read_excel(input$data$datapath),
                 xls = read_excel(input$data$datapath),
                 sav = read_sav(input$data$datapath),
                 dta = read_dta(input$data$datapath),
                 stop("Unsupported file type.")
    )
    
    updateSelectInput(session, "outcome", choices = names(df))
    updateSelectizeInput(session, "predictors", choices = names(df), server = TRUE)
    return(df)
  })
  
  output$outcome_warning <- renderUI({
    req(input$outcome, input$data)
    df <- read_data()
    if (!input$outcome %in% names(df)) return(NULL)
    
    vals <- unique(na.omit(as.character(df[[input$outcome]])))
    if (length(vals) > 2) {
      HTML("<div class='warning-text'><i class='fa fa-exclamation-triangle'></i> Warning: Outcome has >2 levels. Please select a binary variable.</div>")
    } else {
      NULL
    }
  })
  
  output$recode_ui <- renderUI({
    req(input$outcome, input$data)
    df <- read_data()
    choices <- unique(as.character(na.omit(df[[input$outcome]])))
    
    tagList(
      selectInput("old_val1", "Original Value 1", choices = choices, selected = choices[1], width = "100%"),
      textInput("new_val1", "Recode to (should be 0)", value = "0", width = "100%"),
      selectInput("old_val2", "Original Value 2", choices = choices, selected = ifelse(length(choices) > 1, choices[2], choices[1]), width = "100%"),
      textInput("new_val2", "Recode to (should be 1)", value = "1", width = "100%")
    )
  })
  
  analysis_result <- eventReactive(input$generate, {
    tryCatch({
      df <- read_data()
      req(input$outcome, input$predictors)
      
      if (!is.null(input$old_val1) && input$new_val1 != "" &&
          !is.null(input$old_val2) && input$new_val2 != "") {
        df[[input$outcome]][df[[input$outcome]] == input$old_val1] <- input$new_val1
        df[[input$outcome]][df[[input$outcome]] == input$old_val2] <- input$new_val2
      }
      
      all_vars <- c(input$outcome, input$predictors)
      df <- df[, all_vars]
      df <- na.omit(df)
      
      df[[input$outcome]] <- as.numeric(as.character(df[[input$outcome]]))
      if (!all(df[[input$outcome]] %in% c(0, 1))) stop("Outcome must be binary (0 and 1).")
      df[[input$outcome]] <- factor(df[[input$outcome]])
      
      formula <- as.formula(paste(input$outcome, "~", paste(input$predictors, collapse = "+")))
      model <- glm(formula, data = df, family = "binomial")
      
      probs <- predict(model, type = "response")
      pred_class <- factor(ifelse(probs >= 0.5, "1", "0"), levels = c("0", "1"))
      true_class <- factor(df[[input$outcome]], levels = c("0", "1"))
      
      keep <- complete.cases(probs, true_class)
      probs <- probs[keep]
      pred_class <- pred_class[keep]
      true_class <- true_class[keep]
      
      roc_obj <- roc(true_class, probs)
      auc_val <- auc(roc_obj)
      accuracy <- Accuracy(y_pred = pred_class, y_true = true_class)
      conf_mat <- confusionMatrix(pred_class, true_class)
      
      roc_df <- data.frame(
        TPR = rev(roc_obj$sensitivities),
        FPR = rev(1 - roc_obj$specificities)
      )
      
      p <- ggplot(roc_df, aes(x = FPR, y = TPR)) +
        geom_line(color = "#3498db", size = 1.4) +
        geom_ribbon(aes(ymin = 0, ymax = TPR), fill = "#3498db", alpha = 0.1) +
        geom_abline(linetype = "dashed", color = "gray") +
        labs(title = input$graph_title, 
             x = input$x_label, 
             y = input$y_label,
             caption = paste("AUC =", round(auc_val, 3))) +
        theme_minimal(base_size = 16) +
        theme(plot.title = element_text(hjust = 0.5, face = "bold", size = 18),
              axis.title = element_text(face = "bold"),
              panel.grid.minor = element_blank(),
              panel.grid.major = element_line(color = "#f0f0f0"),
              plot.caption = element_text(face = "bold", size = 12))
      
      list(
        roc_obj = roc_obj,
        auc = auc_val,
        accuracy = accuracy,
        conf_mat = conf_mat,
        plot = p
      )
    }, error = function(e) {
      list(error = TRUE, message = e$message)
    })
  })
  
  output$roc_plot <- renderPlotly({
    res <- analysis_result()
    if (!is.null(res$error)) {
      return(plotly_empty(type = "scatter", mode = "lines") %>%
               layout(title = list(text = paste("Error:", res$message), 
                                   font = list(color = "#e74c3c")),
                      plot_bgcolor = "rgba(0,0,0,0)",
                      paper_bgcolor = "rgba(0,0,0,0)"))
    }
    
    ggplotly(res$plot, tooltip = c("x", "y")) %>% 
      layout(hoverlabel = list(bgcolor = "white", 
                               font = list(color = "black")),
             margin = list(t = 60)) %>%
      config(displayModeBar = TRUE)
  })
  
  output$metrics_table <- renderDT({
    res <- analysis_result()
    if (!is.null(res$error)) {
      return(datatable(data.frame(Metric = "Error", Value = res$message), 
                       options = list(dom = 't'), 
                       rownames = FALSE) %>%
               formatStyle(columns = c("Metric", "Value"), 
                           color = "#e74c3c",
                           fontWeight = 'bold'))
    }
    
    metrics_df <- data.frame(
      Metric = c("Accuracy", "AUC", "Sensitivity", "Specificity"),
      Value = c(
        round(res$accuracy, 4),
        round(as.numeric(res$auc), 4),
        round(res$conf_mat$byClass["Sensitivity"], 4),
        round(res$conf_mat$byClass["Specificity"], 4)
      )
    )
    datatable(metrics_df, 
              options = list(dom = 't', ordering = FALSE), 
              rownames = FALSE) %>%
      formatStyle('Metric', fontWeight = 'bold') %>%
      formatStyle('Value', color = '#3498db', fontWeight = 'bold')
  })
  
  output$conf_matrix_table <- renderDT({
    res <- analysis_result()
    if (!is.null(res$error)) return(NULL)
    
    cm <- as.table(res$conf_mat$table)
    cm_df <- as.data.frame.matrix(cm)
    datatable(cm_df, 
              options = list(dom = 't'), 
              rownames = TRUE) %>%
      formatStyle(names(cm_df), fontWeight = 'bold') %>%
      formatStyle(colnames(cm_df)[1], color = '#e74c3c') %>%
      formatStyle(colnames(cm_df)[2], color = '#2ecc71')
  })
  
  output$download_plot <- downloadHandler(
    filename = function() {
      paste0("ROC_Curve_", Sys.Date(), ".png")
    },
    content = function(file) {
      res <- analysis_result()
      if (!is.null(res$error)) stop(res$message)
      ggsave(file, plot = res$plot, device = "png", width = 9, height = 6, dpi = 300)
    }
  )
  
  output$download_word <- downloadHandler(
    filename = function() {
      paste0("ROC_Report_", Sys.Date(), ".docx")
    },
    content = function(file) {
      res <- analysis_result()
      if (!is.null(res$error)) stop(res$message)
      
      tmp_img <- tempfile(fileext = ".png")
      ggsave(tmp_img, plot = res$plot, width = 7, height = 5, dpi = 300)
      
      doc <- read_docx() %>%
        body_add_par("ROC Curve Analysis Report", style = "heading 1") %>%
        body_add_par(paste("Generated on", format(Sys.Date(), "%B %d, %Y")), style = "Normal") %>%
        body_add_break() %>%
        body_add_img(tmp_img, width = 6, height = 4.5) %>%
        body_add_par("Performance Metrics", style = "heading 2") %>%
        body_add_flextable(flextable(data.frame(
          Metric = c("Accuracy", "AUC", "Sensitivity", "Specificity"),
          Value = c(
            round(res$accuracy, 4),
            round(as.numeric(res$auc), 4),
            round(res$conf_mat$byClass["Sensitivity"], 4),
            round(res$conf_mat$byClass["Specificity"], 4)
          )
        ) %>% 
          theme_box() %>%
          set_table_properties(layout = "autofit")) %>%
          body_add_par("Confusion Matrix", style = "heading 2") %>%
          body_add_flextable(flextable(as.data.frame.matrix(as.table(res$conf_mat$table))) %>%
                               theme_box() %>%
                               set_table_properties(layout = "autofit")) %>%
          body_add_par("Analysis Details", style = "heading 2") %>%
          body_add_par(paste("Outcome variable:", input$outcome)), style = "Normal") %>%
        body_add_par(paste("Predictor variables:", paste(input$predictors, collapse = ", ")), style = "Normal") %>%
        body_add_par(paste("Observations:", nrow(na.omit(read_data()[, c(input$outcome, input$predictors)]))), style = "Normal") %>%
        body_add_break() %>%
        body_add_par("ROC Curve Builder", style = "Normal") %>%
        body_add_par("For research and educational use", style = "subtle")
      
      print(doc, target = file)
    }
  )
}

shinyApp(ui, server)
