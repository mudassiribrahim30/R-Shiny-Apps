library(shiny)
library(readxl)
library(haven)
library(officer)
library(flextable)
library(writexl)
library(shinythemes)
library(dplyr)
library(DT)
library(tools)
library(shinyjs)

ui <- fluidPage(
  theme = shinytheme("flatly"),
  useShinyjs(),
  tags$head(
    tags$style(HTML("
      .sidebar-panel {
        background-color: #f8f9fa;
        border-radius: 5px;
        padding: 20px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
      }
      .btn-primary {
        background-color: #3498db;
        border-color: #3498db;
      }
      .btn-success {
        background-color: #2ecc71;
        border-color: #2ecc71;
      }
      .btn-info {
        background-color: #17a2b8;
        border-color: #17a2b8;
      }
      .footer {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 5px;
        margin-top: 20px;
        font-size: 14px;
        color: #7f8c8d;
      }
      .info-box {
        background-color: #e8f4fc;
        border-left: 5px solid #3498db;
        padding: 15px;
        margin-bottom: 20px;
        border-radius: 0 5px 5px 0;
      }
      .app-title {
        color: #2c3e50;
        margin-bottom: 5px;
      }
      .app-subtitle {
        color: #7f8c8d;
        font-size: 16px;
        margin-bottom: 20px;
      }
      .progress-bar {
        height: 10px;
        margin-top: 5px;
      }
      .large-file-warning {
        color: #e74c3c;
        font-weight: bold;
      }
    "))
  ),
  
  titlePanel(
    div(
      h1(class = "app-title", "TagSelect: High-Capacity Random Sampler"),
      h4(class = "app-subtitle", "Efficient participant (s) selection for small and large research studies")
    )
  ),
  
  sidebarLayout(
    sidebarPanel(
      class = "sidebar-panel",
      fileInput("file", "Upload Participant (s) List", 
                accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta", ".docx"),
                buttonLabel = "Browse...",
                placeholder = "No file selected"),
      div(id = "file_stats",
          textOutput("file_size"),
          textOutput("dimensions"),
          tags$small("Supports up to 10,000 participants and 1,000 variables", 
                    class = "text-muted")
      ),
      br(),
      numericInput("sample_size", "Number of Participants to Sample:", 
                  value = 5, min = 1, max = 10000, step = 1),
      actionButton("sample_btn", "Select Random Sample", 
                  class = "btn-primary", icon = icon("random")),
      br(), br(),
      
      div(class = "info-box",
          h4(icon("info-circle"), "Instructions"),
          tags$ol(
            tags$li("Upload your list of participants (up to 10,000 rows and 1,000 columns)"),
            tags$li("Enter how many participants you want to randomly select"),
            tags$li("Click 'Select Random Sample' to tag the selected participants"),
            tags$li("Selected participants will be marked with 'Selected = TRUE'"),
            tags$li("Download either the full list or only the selected participants")
          )
      ),
      
      h4(icon("cogs"), "Technical Details"),
      helpText("Optimized for large datasets using efficient algorithms:"),
      code("sample.int(n, size)"),
      helpText("This uses Fisher-Yates shuffling for better performance with large n"),
      br(),
      
      h4(icon("download"), "Download Options"),
      div(style = "margin-bottom: 10px;",
          downloadButton("download_all_excel", "Download Full List (Excel)", 
                        class = "btn-success", icon = icon("file-excel"))
      ),
      div(style = "margin-bottom: 10px;",
          downloadButton("download_selected_excel", "Download Selected Only (Excel)", 
                        class = "btn-success", icon = icon("file-excel"))
      ),
      div(style = "margin-bottom: 10px;",
          downloadButton("download_all_word", "Download Full List (Word)", 
                        class = "btn-info", icon = icon("file-word"))
      ),
      downloadButton("download_selected_word", "Download Selected Only (Word)", 
                    class = "btn-info", icon = icon("file-word"))
    ),

    mainPanel(
      h3(icon("users"), "Participant (s) List"),
      textOutput("obs_count"),
      conditionalPanel(
        condition = "input.file != null",
        div(id = "progress_container",
            shiny::tags$div(class = "progress",
                shiny::tags$div(id = "progress_bar", 
                                class = "progress-bar progress-bar-striped active",
                                role = "progressbar",
                                style = "width: 0%"))
        )
      ),
      br(),
      DTOutput("table"),
      
      div(class = "footer",
          HTML("<strong>Developer:</strong> Mudasir Mohammed Ibrahim | 
               <strong>Contact:</strong> <a href='mailto:mudassiribrahim30@gmail.com'>mudassiribrahim30@gmail.com</a><br>
               <strong>Version:</strong> 2.1 | 
               <strong>Capacity:</strong> 10,000 participants × 1,000 variables"),
          tags$br(),
          tags$a(href = "https://github.com/mudasir-ibrahim/TagSelect", 
                 target = "_blank", 
                 icon("github"), "View on GitHub")
      )
    )
  )
)

server <- function(input, output, session) {
  data <- reactiveVal(NULL)
  
  # Initialize progress bar
  observe({
    if (is.null(input$file)) {
      shinyjs::hide("progress_container")
    } else {
      shinyjs::show("progress_container")
    }
  })
  
  # Update progress bar
  update_progress <- function(value) {
    shinyjs::runjs(sprintf('$("#progress_bar").css("width", "%s%%")', value))
  }
  
  output$file_size <- renderText({
    req(input$file)
    size <- file.size(input$file$datapath)
    paste("File size:", format(structure(size, class = "object_size"), 
          units = "auto"))
  })
  
  output$dimensions <- renderText({
    req(data())
    paste("Dimensions:", nrow(data()), "participants ×", ncol(data()), "variables")
  })
  
  observeEvent(input$file, {
    req(input$file)
    ext <- tools::file_ext(input$file$name)
    
    tryCatch({
      update_progress(10)
      
      # Read data with progress updates
      df <- switch(ext,
        csv = {
          update_progress(20)
          read.csv(input$file$datapath, stringsAsFactors = FALSE)
        },
        xlsx = {
          update_progress(20)
          read_excel(input$file$datapath)
        },
        xls = {
          update_progress(20)
          read_excel(input$file$datapath)
        },
        sav = {
          update_progress(20)
          read_sav(input$file$datapath)
        },
        dta = {
          update_progress(20)
          read_dta(input$file$datapath)
        },
        docx = {
          update_progress(20)
          tmp <- officer::read_docx(input$file$datapath)
          tables <- officer::docx_summary(tmp)
          table_data <- tables[tables$content_type == "table cell", ]
          wide_data <- reshape(table_data[, c("row_id", "cell_id", "text")],
                               timevar = "cell_id",
                               idvar = "row_id",
                               direction = "wide")
          names(wide_data) <- make.names(paste0("V", seq_along(wide_data)))
          wide_data
        },
        validate("Unsupported file type.")
      )
      
      update_progress(70)
      
      # Check dimensions
      if (nrow(df) > 10000) {
        showNotification("Warning: Dataset exceeds 10,000 participants. Only first 10,000 will be used.", 
                        type = "warning")
        df <- df[1:10000, ]
      }
      
      if (ncol(df) > 1000) {
        showNotification("Warning: Dataset exceeds 1,000 variables. Only first 1,000 will be used.", 
                        type = "warning")
        df <- df[, 1:1000]
      }
      
      df$Selected <- FALSE
      data(df)
      update_progress(100)
      
      showNotification("File successfully uploaded!", type = "message")
    }, error = function(e) {
      update_progress(0)
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  observeEvent(input$sample_btn, {
    df <- data()
    req(df)
    
    sample_size <- input$sample_size
    total <- nrow(df)
    
    if (sample_size > total) {
      showNotification("Sample size exceeds number of participants.", type = "error")
      return()
    }
    
    # Use more efficient sampling for large datasets
    df$Selected <- FALSE
    sample_rows <- sample.int(total, sample_size)
    df$Selected[sample_rows] <- TRUE
    data(df)
    
    showNotification(paste(sample_size, "participants randomly selected!"), 
                    type = "message")
  })
  
  output$table <- renderDT({
    req(data())
    datatable(
      data(),
      options = list(
        pageLength = 10,
        scrollX = TRUE,
        scrollY = "500px",
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel', 'pdf'),
        deferRender = TRUE,  # Better performance for large tables
        scroller = TRUE      # Smooth scrolling
      ),
      rownames = FALSE,
      filter = 'top',
      class = 'hover stripe nowrap',
      selection = 'none'
    ) %>% 
      formatStyle(
        'Selected',
        backgroundColor = styleEqual(c(TRUE, FALSE), c('#d4edda', '#f8d7da'))
      )
  })
  
  output$obs_count <- renderText({
    req(data())
    selected_count <- sum(data()$Selected)
    paste("Total participants:", nrow(data()), "| Selected:", selected_count)
  })
  
  # Excel downloads
  output$download_all_excel <- downloadHandler(
    filename = function() { 
      paste0("FullList_Tagged_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx") 
    },
    content = function(file) {
      withProgress(
        message = 'Preparing Excel file',
        detail = 'This may take a moment...',
        value = 0,
        {
          incProgress(0.3)
          write_xlsx(data(), path = file)
          incProgress(1)
        }
      )
    }
  )
  
  output$download_selected_excel <- downloadHandler(
    filename = function() { 
      paste0("SelectedParticipants_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx") 
    },
    content = function(file) {
      selected_data <- data() %>% filter(Selected == TRUE)
      withProgress(
        message = 'Preparing Excel file',
        detail = 'This may take a moment...',
        value = 0,
        {
          incProgress(0.3)
          write_xlsx(selected_data, path = file)
          incProgress(1)
        }
      )
    }
  )
  
  # Word downloads
  output$download_all_word <- downloadHandler(
    filename = function() { 
      paste0("FullList_Tagged_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".docx") 
    },
    content = function(file) {
      withProgress(
        message = 'Preparing Word document',
        detail = 'This may take a moment...',
        value = 0,
        {
          incProgress(0.2)
          doc <- read_docx()
          incProgress(0.3)
          doc <- body_add_par(doc, "Full List of Participants (Tagged)", style = "heading 1")
          doc <- body_add_par(doc, paste("Generated on:", Sys.Date()), style = "Normal")
          doc <- body_add_par(doc, paste("Total participants:", nrow(data())), style = "Normal")
          incProgress(0.5)
          doc <- body_add_flextable(doc, flextable(data()) %>% 
                                    theme_zebra() %>%
                                    bg(i = ~ Selected == TRUE, bg = "#d4edda"))
          incProgress(0.8)
          print(doc, target = file)
          incProgress(1)
        }
      )
    }
  )
  
  output$download_selected_word <- downloadHandler(
    filename = function() { 
      paste0("SelectedParticipants_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".docx") 
    },
    content = function(file) {
      selected_data <- data() %>% filter(Selected == TRUE)
      withProgress(
        message = 'Preparing Word document',
        detail = 'This may take a moment...',
        value = 0,
        {
          incProgress(0.2)
          doc <- read_docx()
          incProgress(0.3)
          doc <- body_add_par(doc, "Selected Participants Only", style = "heading 1")
          doc <- body_add_par(doc, paste("Generated on:", Sys.Date()), style = "Normal")
          doc <- body_add_par(doc, paste("Total selected:", nrow(selected_data)), style = "Normal")
          incProgress(0.5)
          doc <- body_add_flextable(doc, flextable(selected_data) %>% theme_zebra())
          incProgress(0.8)
          print(doc, target = file)
          incProgress(1)
        }
      )
    }
  )
}

shinyApp(ui, server)
