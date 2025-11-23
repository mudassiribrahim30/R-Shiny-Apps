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
library(shinyWidgets)
library(randomizr)
library(tidyr)
library(purrr)

ui <- fluidPage(
  useShinyjs(),
  tags$head(
    tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css"),
    tags$script(HTML("
      // Function to toggle theme
      function toggleTheme() {
        const body = document.body;
        const currentTheme = body.getAttribute('data-theme');
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        body.setAttribute('data-theme', newTheme);
        
        // Store theme preference in localStorage
        localStorage.setItem('appTheme', newTheme);
        
        // Update button icon
        const themeIcon = document.getElementById('theme-icon');
        if (newTheme === 'dark') {
          themeIcon.className = 'fas fa-sun';
          themeIcon.style.color = '#ffd43b';
        } else {
          themeIcon.className = 'fas fa-moon';
          themeIcon.style.color = '#f8f9fa';
        }
      }
      
      // Apply saved theme on page load
      document.addEventListener('DOMContentLoaded', function() {
        const savedTheme = localStorage.getItem('appTheme') || 'light';
        document.body.setAttribute('data-theme', savedTheme);
        
        // Update button icon based on saved theme
        const themeIcon = document.getElementById('theme-icon');
        if (savedTheme === 'dark') {
          themeIcon.className = 'fas fa-sun';
          themeIcon.style.color = '#ffd43b';
        } else {
          themeIcon.className = 'fas fa-moon';
          themeIcon.style.color = '#f8f9fa';
        }
      });
    ")),
    tags$style(HTML("
      /* Light theme (default) */
      :root {
        --bg-primary: white;
        --bg-secondary: #f8f9fa;
        --bg-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        --text-primary: #2c3e50;
        --text-secondary: #7f8c8d;
        --text-muted: #6c757d;
        --text-heading: #2c3e50;
        --text-body: #495057;
        --card-bg: white;
        --card-shadow: 0 4px 6px rgba(0,0,0,0.1);
        --border-color: #e9ecef;
        --success-color: #27ae60;
        --info-color: #17a2b8;
        --warning-color: #f39c12;
        --danger-color: #e74c3c;
        --primary-gradient: linear-gradient(135deg, #667eea, #764ba2);
      }
      
      /* Dark theme */
      [data-theme='dark'] {
        --bg-primary: #0f1419;
        --bg-secondary: #1e252b;
        --bg-gradient: linear-gradient(135deg, #1a365d 0%, #2d3748 100%);
        --text-primary: #e2e8f0;
        --text-secondary: #cbd5e0;
        --text-muted: #a0aec0;
        --text-heading: #f7fafc;
        --text-body: #e2e8f0;
        --card-bg: #1e252b;
        --card-shadow: 0 4px 6px rgba(0,0,0,0.3);
        --border-color: #2d3748;
        --success-color: #48bb78;
        --info-color: #4299e1;
        --warning-color: #ed8936;
        --danger-color: #f56565;
        --primary-gradient: linear-gradient(135deg, #4a5568, #2d3748);
      }
      
      @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;600;700&display=swap');
      
      * {
        font-family: 'Roboto', sans-serif;
        transition: background-color 0.3s ease, color 0.3s ease, border-color 0.3s ease;
      }
      
      body {
        background: var(--bg-gradient);
        min-height: 100vh;
        margin: 0;
        padding: 0;
        color: var(--text-body);
      }
      
      .main-container {
        background: var(--bg-primary);
        border-radius: 0px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        padding: 0;
        min-height: 100vh;
        display: flex;
        flex-direction: column;
        width: 100%;
        margin: 0;
      }
      
      @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
      }
      
      @keyframes slideIn {
        from { transform: translateY(-30px); opacity: 0; }
        to { transform: translateY(0); opacity: 1; }
      }
      
      .title-animation {
        animation: slideIn 1s ease-out;
        text-align: center;
        color: var(--text-heading);
        font-weight: 700;
        margin-bottom: 10px;
        background: var(--primary-gradient);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.5em;
      }
      
      .subtitle-animation {
        animation: fadeIn 1.5s ease-out;
        text-align: center;
        color: var(--text-secondary);
        margin-bottom: 30px;
        font-weight: 400;
        font-size: 1.2em;
      }
      
      .developer-info {
        background: var(--primary-gradient);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 25px;
        text-align: center;
        animation: fadeIn 2s ease-out;
      }
      
      .formula-box {
        background: var(--bg-secondary);
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
        border-left: 5px solid var(--info-color);
        box-shadow: var(--card-shadow);
        color: var(--text-body);
      }
      
      .result-box {
        background: var(--card-bg);
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
        border-left: 5px solid var(--success-color);
        box-shadow: var(--card-shadow);
        color: var(--text-body);
      }
      
      .info-box {
        background: var(--card-bg);
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
        border-left: 5px solid var(--warning-color);
        box-shadow: var(--card-shadow);
        color: var(--text-body);
      }
      
      .info-box h4, .formula-box h4, .result-box h4, .calculation-box h4 {
        color: var(--text-heading) !important;
        font-weight: 600;
      }
      
      .calculation-box {
        background: var(--card-bg);
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
        border-left: 5px solid var(--info-color);
        box-shadow: var(--card-shadow);
        color: var(--text-body);
      }
      
      .random-start-box {
        background: var(--card-bg);
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
        border-left: 5px solid #9b59b6;
        box-shadow: var(--card-shadow);
        animation: fadeIn 0.5s ease-out;
        color: var(--text-body);
      }
      
      .navbar-default .navbar-brand {
        color: white !important;
        font-weight: 700;
        font-size: 24px;
        background: var(--primary-gradient);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        padding: 15px;
        margin-left: 10px;
      }
      
      .nav-tabs > li > a {
        color: var(--text-primary);
        font-weight: 500;
        border-radius: 8px 8px 0 0;
        margin-right: 5px;
        padding: 12px 20px;
        background: var(--bg-secondary);
      }
      
      .nav-tabs > li.active > a {
        background: var(--primary-gradient) !important;
        color: white !important;
        border: none;
      }
      
      .btn-primary {
        background: var(--primary-gradient);
        border: none;
        border-radius: 8px;
        font-weight: 500;
        padding: 10px 20px;
        transition: all 0.3s ease;
        color: white;
      }
      
      .btn-primary:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        color: white;
      }
      
      .btn-warning {
        background: linear-gradient(135deg, var(--warning-color), var(--danger-color));
        border: none;
        border-radius: 8px;
        font-weight: 500;
        padding: 10px 20px;
        transition: all 0.3s ease;
        color: white;
      }
      
      .btn-success {
        background: linear-gradient(135deg, var(--success-color), #2ecc71);
        border: none;
        border-radius: 8px;
        font-weight: 500;
        padding: 10px 20px;
        transition: all 0.3s ease;
        color: white;
      }
      
      .sidebar-panel {
        background: var(--bg-secondary);
        padding: 25px;
        border-radius: 10px;
        box-shadow: var(--card-shadow);
        height: fit-content;
        color: var(--text-body);
      }
      
      .main-panel {
        padding: 0 25px;
        width: 100%;
        background: var(--bg-primary);
        color: var(--text-body);
      }
      
      .download-btn {
        width: 100%; 
        margin-bottom: 10px; 
        background: var(--primary-gradient); 
        color: white; 
        border: none; 
        border-radius: 8px; 
        padding: 10px;
      }
      
      .footer {
        background: var(--primary-gradient);
        color: white;
        text-align: center;
        padding: 15px;
        border-radius: 0px;
        margin-top: auto;
        font-size: 0.9em;
        width: 100%;
      }
      
      .content-wrapper {
        flex: 1;
        display: flex;
        flex-direction: column;
        width: 100%;
        padding: 20px;
        background: var(--bg-primary);
        color: var(--text-body);
      }
      
      .main-content-area {
        flex: 1;
        margin-bottom: 20px;
        width: 100%;
        background: var(--bg-primary);
        color: var(--text-body);
      }
      
      .plot-download-btn {
        background: linear-gradient(135deg, var(--success-color), #2ecc71);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        margin: 10px 5px;
        font-weight: 500;
        transition: all 0.3s ease;
      }
      
      .plot-download-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(39, 174, 96, 0.4);
        color: white;
      }
      
      .sidebar-panel {
        position: relative;
        height: auto;
        max-height: none;
        overflow-y: visible;
      }
      
      .main-panel {
        height: auto;
        max-height: none;
        overflow-y: visible;
      }
      
      .file-upload-info {
        background: var(--bg-secondary);
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
        font-size: 0.9em;
        border-left: 3px solid var(--info-color);
        color: var(--text-body);
      }
      
      .status-card {
        background: var(--card-bg);
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 15px;
        box-shadow: var(--card-shadow);
        border-left: 4px solid var(--info-color);
        color: var(--text-body);
      }
      
      .status-title {
        font-size: 14px;
        color: var(--text-secondary);
        margin-bottom: 5px;
        font-weight: 500;
      }
      
      .status-value {
        font-size: 18px;
        font-weight: 600;
        color: var(--text-primary);
      }
      
      .section-title {
        color: var(--text-heading);
        border-bottom: 2px solid var(--border-color);
        padding-bottom: 8px;
        margin-bottom: 15px;
        font-weight: 600;
        font-size: 1.3em;
      }
      
      .download-section {
        background-color: var(--bg-secondary);
        border-radius: 8px;
        padding: 15px;
        margin-top: 20px;
        border: 1px dashed var(--border-color);
        color: var(--text-body);
      }
      
      .download-title {
        font-size: 16px;
        font-weight: 600;
        color: var(--text-heading);
        margin-bottom: 15px;
      }
      
      .action-btn {
        width: 100%;
        margin-bottom: 10px;
      }
      
      .data-table {
        border-radius: 8px;
        overflow: hidden;
        width: 100% !important;
        background: var(--card-bg);
        color: var(--text-body);
      }
      
      .selected-row {
        background-color: rgba(39, 174, 96, 0.2) !important;
      }
      
      .randomization-card {
        background-color: var(--card-bg);
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: var(--card-shadow);
        border-left: 4px solid var(--info-color);
        color: var(--text-body);
      }
      
      .randomization-header {
        font-size: 18px;
        font-weight: 600;
        color: var(--text-heading);
        margin-bottom: 15px;
      }
      
      .randomization-description {
        font-size: 14px;
        color: var(--text-secondary);
        margin-bottom: 15px;
      }
      
      .group-allocation-box {
        background-color: var(--bg-secondary);
        border-radius: 8px;
        padding: 15px;
        margin-top: 15px;
        border: 1px solid var(--border-color);
        color: var(--text-body);
      }
      
      .group-allocation-row {
        margin-bottom: 10px;
      }
      
      .file-input-label {
        font-weight: 500;
        margin-bottom: 8px;
        color: var(--text-primary);
      }
      
      .numeric-input-label {
        font-weight: 500;
        margin-bottom: 8px;
        color: var(--text-primary);
      }
      
      .large-file-warning {
        color: var(--danger-color);
        font-weight: bold;
      }
      
      .progress-bar {
        height: 10px;
        margin-top: 5px;
        border-radius: 4px;
      }
      
      /* Theme toggle button - FIXED POSITION */
      .theme-toggle {
        position: fixed;
        top: 20px;
        right: 30px;
        background: var(--primary-gradient);
        border: none;
        border-radius: 50%;
        width: 50px;
        height: 50px;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 20px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.3);
        transition: all 0.3s ease;
        z-index: 9999;
        border: 2px solid rgba(255,255,255,0.2);
      }
      
      .theme-toggle:hover {
        transform: scale(1.15) rotate(15deg);
        box-shadow: 0 8px 16px rgba(0,0,0,0.4);
      }
      
      /* Fix for navbar brand */
      .navbar-brand {
        padding: 15px 15px !important;
        height: auto !important;
        font-size: 24px !important;
        line-height: 1.5 !important;
      }
      
      /* Ensure full width layout */
      .container-fluid {
        padding-left: 0;
        padding-right: 0;
        width: 100%;
        margin: 0;
      }
      
      .row {
        margin-left: 0;
        margin-right: 0;
        width: 100%;
      }
      
      .col-sm-4, .col-sm-8, .col-lg-4, .col-lg-8, .col-md-4, .col-md-8 {
        padding-left: 15px;
        padding-right: 15px;
      }
      
      /* Fix tab content width */
      .tab-content {
        width: 100%;
        background: var(--bg-primary);
        color: var(--text-body);
      }
      
      /* Ensure proper spacing in main containers */
      .main-content-area .container-fluid {
        padding: 20px;
        background: var(--bg-primary);
        color: var(--text-body);
      }
      
      /* Text styling improvements for dark theme */
      [data-theme='dark'] {
        color: var(--text-body);
      }
      
      [data-theme='dark'] h1, 
      [data-theme='dark'] h2, 
      [data-theme='dark'] h3, 
      [data-theme='dark'] h4, 
      [data-theme='dark'] h5, 
      [data-theme='dark'] h6 {
        color: var(--text-heading) !important;
        font-weight: 600;
      }
      
      [data-theme='dark'] p {
        color: var(--text-body) !important;
        font-weight: 400;
        line-height: 1.6;
      }
      
      [data-theme='dark'] li {
        color: var(--text-body) !important;
        font-weight: 400;
      }
      
      [data-theme='dark'] strong {
        color: var(--text-heading) !important;
        font-weight: 600;
      }
      
      [data-theme='dark'] .formula-box p,
      [data-theme='dark'] .info-box p,
      [data-theme='dark'] .result-box p,
      [data-theme='dark'] .calculation-box p {
        color: var(--text-body) !important;
      }
      
      /* Dark theme specific adjustments for data tables */
      [data-theme='dark'] .dataTables_wrapper .dataTables_length,
      [data-theme='dark'] .dataTables_wrapper .dataTables_filter,
      [data-theme='dark'] .dataTables_wrapper .dataTables_info,
      [data-theme='dark'] .dataTables_wrapper .dataTables_processing,
      [data-theme='dark'] .dataTables_wrapper .dataTables_paginate {
        color: var(--text-primary) !important;
      }
      
      [data-theme='dark'] table.dataTable thead th,
      [data-theme='dark'] table.dataTable thead td {
        border-bottom: 1px solid var(--border-color);
        color: var(--text-heading);
        background: var(--bg-secondary);
        font-weight: 600;
      }
      
      [data-theme='dark'] table.dataTable tbody th,
      [data-theme='dark'] table.dataTable tbody td {
        border-bottom: 1px solid var(--border-color);
        color: var(--text-body);
        background: var(--card-bg);
      }
      
      [data-theme='dark'] .dataTables_wrapper .dataTables_paginate .paginate_button {
        color: var(--text-primary) !important;
      }
      
      /* Form controls in dark theme */
      [data-theme='dark'] .form-control {
        background-color: var(--bg-secondary);
        border-color: var(--border-color);
        color: var(--text-body);
      }
      
      [data-theme='dark'] .form-control:focus {
        background-color: var(--bg-secondary);
        border-color: var(--info-color);
        color: var(--text-body);
        box-shadow: 0 0 0 0.2rem rgba(66, 153, 225, 0.25);
      }
      
      [data-theme='dark'] .form-control::placeholder {
        color: var(--text-muted);
      }
      
      /* Dropdowns in dark theme */
      [data-theme='dark'] .dropdown-menu {
        background-color: var(--card-bg);
        border-color: var(--border-color);
      }
      
      [data-theme='dark'] .dropdown-menu > li > a {
        color: var(--text-body);
      }
      
      [data-theme='dark'] .dropdown-menu > li > a:hover {
        background-color: var(--bg-secondary);
        color: var(--text-heading);
      }
      
      /* Radio buttons and checkboxes in dark theme */
      [data-theme='dark'] .radio label,
      [data-theme='dark'] .checkbox label {
        color: var(--text-body);
      }
      
      /* Well panels in dark theme */
      [data-theme='dark'] .well {
        background-color: var(--bg-secondary);
        border-color: var(--border-color);
        color: var(--text-body);
      }
      
      /* HR styling in dark theme */
      [data-theme='dark'] hr {
        border-color: var(--border-color);
      }
      
      /* Improved contrast for better readability */
      [data-theme='dark'] {
        --text-heading: #f7fafc;
        --text-body: #e2e8f0;
        --text-secondary: #cbd5e0;
        --text-muted: #a0aec0;
      }
      
      /* Additional text contrast improvements */
      [data-theme='dark'] .sidebar-panel,
      [data-theme='dark'] .main-panel,
      [data-theme='dark'] .tab-content {
        color: var(--text-body);
      }
      
      [data-theme='dark'] .section-title {
        color: var(--text-heading);
        font-weight: 600;
      }
      
      [data-theme='dark'] .file-input-label,
      [data-theme='dark'] .numeric-input-label {
        color: var(--text-primary);
        font-weight: 500;
      }
    "))
  ),
  
  # Theme toggle button - FIXED POSITION (now in front of everything)
  div(class = "theme-toggle", 
      onclick = "toggleTheme()",
      title = "Toggle Dark/Light Mode",
      icon("moon", id = "theme-icon")
  ),
  
  div(class = "main-container",
      div(class = "content-wrapper",
          navbarPage(
            title = div(
              style = "color: white; font-weight: 700; font-size: 24px; background: var(--primary-gradient); -webkit-background-clip: text; -webkit-text-fill-color: transparent; padding: 5px 0;",
              "TagSelect"
            ),
            windowTitle = "TagSelect: High-Capacity Random Sampler",
            collapsible = TRUE,
            inverse = FALSE,
            fluid = TRUE,
            header = tags$head(
              tags$style(HTML("
                .navbar {
                  border-radius: 0;
                  margin-bottom: 0;
                  background-color: var(--bg-primary);
                  border: none;
                  box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                  position: relative;
                  z-index: 1000;
                }
                .navbar-default .navbar-nav > li > a {
                  color: var(--text-primary);
                  font-weight: 500;
                }
                .navbar-default .navbar-nav > .active > a {
                  background-color: transparent !important;
                  color: #667eea !important;
                }
              "))
            ),
            
            tabPanel(
              "Simple Randomization",
              div(
                class = "container-fluid",
                style = "padding: 20px; width: 100%;",
                
                fluidRow(
                  column(
                    12,
                    h1(class = "title-animation", "Simple Randomization"),
                    h4(class = "subtitle-animation", "Efficient participant selection for research studies of all sizes"),
                    hr(style = "border-top:1px solid var(--border-color);")
                  )
                ),
                
                fluidRow(
                  column(
                    4,
                    div(
                      class = "sidebar-panel",
                      
                      h4(class = "section-title", icon("upload", class = "fa-fw"), "Data Upload"),
                      
                      div(class = "file-input-label", "Upload Participant List"),
                      fileInput("file", NULL,
                                accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta", ".docx"),
                                buttonLabel = "Browse...",
                                placeholder = "No file selected",
                                width = "100%"),
                      
                      div(id = "file_stats",
                          uiOutput("file_info_box")
                      ),
                      
                      br(),
                      
                      div(class = "numeric-input-label", "Number of Participants to Sample"),
                      numericInput("sample_size", NULL,
                                   value = 5, min = 1, max = 10000, step = 1,
                                   width = "100%"),
                      
                      actionButton("sample_btn", "Select Random Sample", 
                                   class = "btn-primary action-btn", 
                                   icon = icon("random")),
                      
                      br(), br(),
                      
                      div(
                        class = "download-section",
                        h4(class = "download-title", icon("download", class = "fa-fw"), "Download Options"),
                        
                        div(
                          style = "margin-bottom:15px;",
                          dropdownButton(
                            label = "Download Full List",
                            status = "success",
                            icon = icon("file-export"),
                            circle = FALSE,
                            size = "sm",
                            downloadButton("download_all_excel", "Excel", class = "btn-success"),
                            downloadButton("download_all_word", "Word", class = "btn-info"),
                            downloadButton("download_all_csv", "CSV", class = "btn-secondary")
                          )
                        ),
                        
                        div(
                          dropdownButton(
                            label = "Download Selected Only",
                            status = "success",
                            icon = icon("file-export"),
                            circle = FALSE,
                            size = "sm",
                            downloadButton("download_selected_excel", "Excel", class = "btn-success"),
                            downloadButton("download_selected_word", "Word", class = "btn-info"),
                            downloadButton("download_selected_csv", "CSV", class = "btn-secondary")
                          )
                        )
                      ),
                      
                      br(),
                      
                      div(
                        class = "info-box",
                        h4(icon("info-circle", class = "fa-fw"), "Instructions"),
                        tags$ol(
                          tags$li("Upload your list of participants (Excel, CSV, SPSS, Stata, or Word)"),
                          tags$li("Specify how many participants to randomly select"),
                          tags$li("Click 'Select Random Sample' to tag participants"),
                          tags$li("Selected participants will be marked in the table"),
                          tags$li("Download either the full list or only selected participants")
                        )
                      )
                    )
                  ),
                  
                  column(
                    8,
                    div(
                      class = "main-panel",
                      
                      fluidRow(
                        column(
                          6,
                          div(
                            class = "status-card",
                            div(class = "status-title", "Total Participants"),
                            div(class = "status-value", textOutput("total_count"))
                          )
                        ),
                        column(
                          6,
                          div(
                            class = "status-card",
                            div(class = "status-title", "Selected Participants"),
                            div(class = "status-value", textOutput("selected_count"))
                          )
                        )
                      ),
                      
                      conditionalPanel(
                        condition = "input.file != null",
                        div(id = "progress_container",
                            div(class = "progress",
                                div(id = "progress_bar", 
                                    class = "progress-bar progress-bar-striped active",
                                    role = "progressbar",
                                    style = "width: 0%"))
                        )
                      ),
                      
                      br(),
                      
                      h4(class = "section-title", icon("table", class = "fa-fw"), "Participant List"),
                      
                      div(
                        class = "data-table",
                        DTOutput("table")
                      )
                    )
                  )
                )
              )
            ),
            
            tabPanel(
              "Experimental Randomization",
              div(
                class = "container-fluid",
                style = "padding: 20px; width: 100%;",
                
                fluidRow(
                  column(
                    12,
                    h1(class = "title-animation", "Experimental Randomization"),
                    h4(class = "subtitle-animation", "Advanced randomization methods for experimental designs"),
                    hr(style = "border-top:1px solid var(--border-color);")
                  )
                ),
                
                fluidRow(
                  column(
                    4,
                    div(
                      class = "sidebar-panel",
                      
                      h4(class = "section-title", icon("random", class = "fa-fw"), "Randomization Setup"),
                      
                      radioButtons("has_participant_list", "Do you have a list of participants?",
                                   choices = c("Yes" = "yes", "No" = "no"),
                                   selected = "no"),
                      
                      conditionalPanel(
                        condition = "input.has_participant_list == 'yes'",
                        div(
                          class = "info-box",
                          h4(icon("upload", class = "fa-fw"), "Upload Participant List"),
                          fileInput("exp_file", NULL,
                                    accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta", ".docx"),
                                    buttonLabel = "Browse...",
                                    placeholder = "No file selected",
                                    width = "100%"),
                          
                          numericInput("total_participants", "Total Number of Participants:",
                                       value = NULL, min = 1, max = 10000, step = 1),
                          
                          selectInput("group_var", "Select Grouping Variable:", 
                                      choices = c("None" = "none"), selected = "none"),
                          
                          conditionalPanel(
                            condition = "input.group_var != 'none'",
                            checkboxInput("specify_group_n", "Specify sample size for each group?", FALSE),
                            conditionalPanel(
                              condition = "input.specify_group_n",
                              uiOutput("group_allocation_ui")
                            )
                          )
                        )
                      ),
                      
                      conditionalPanel(
                        condition = "input.has_participant_list == 'no'",
                        div(
                          class = "info-box",
                          h4(icon("users", class = "fa-fw"), "Generate Participants"),
                          numericInput("total_n", "Total Number of Participants:",
                                       value = 100, min = 1, max = 10000, step = 1),
                          
                          checkboxInput("has_groups", "Include groups in your design?", FALSE),
                          
                          conditionalPanel(
                            condition = "input.has_groups",
                            textInput("group_names", "Group Names (comma separated):",
                                      placeholder = "e.g., Male, Female"),
                            uiOutput("group_size_ui")
                          )
                        )
                      ),
                      
                      h4(class = "section-title", icon("cogs", class = "fa-fw"), "Randomization Settings"),
                      
                      selectInput("randomization_method", "Select Randomization Method:",
                                  choices = c("Simple Randomization" = "simple",
                                              "Complete Randomization" = "complete",
                                              "Block Randomization" = "block",
                                              "Cluster Randomization" = "cluster",
                                              "Stratified Randomization" = "stratified")),
                      
                      conditionalPanel(
                        condition = "input.randomization_method == 'block'",
                        numericInput("block_size", "Block Size:", value = 4, min = 2, max = 20, step = 1)
                      ),
                      
                      numericInput("num_treatments", "Number of Treatment Groups:", 
                                   value = 2, min = 2, max = 500, step = 1),
                      
                      textInput("treatment_names", "Treatment Group Names (comma separated):",
                                value = "Treatment, Control"),
                      
                      actionButton("randomize_btn", "Perform Randomization", 
                                   class = "btn-primary action-btn", 
                                   icon = icon("random")),
                      
                      br(), br(),
                      
                      div(
                        class = "download-section",
                        h4(class = "download-title", icon("download", class = "fa-fw"), "Download Results"),
                        
                        downloadButton("download_randomized", "Download Randomized Data", 
                                       class = "btn-success action-btn")
                      )
                    )
                  ),
                  
                  column(
                    8,
                    div(
                      class = "main-panel",
                      
                      fluidRow(
                        column(
                          12,
                          div(
                            class = "randomization-card",
                            div(class = "randomization-header", "Randomization Summary"),
                            verbatimTextOutput("randomization_summary")
                          )
                        )
                      ),
                      
                      h4(class = "section-title", icon("table", class = "fa-fw"), "Randomized Data"),
                      
                      div(
                        class = "data-table",
                        DTOutput("randomized_table")
                      )
                    )
                  )
                )
              )
            ),
            
            tabPanel(
              "About",
              div(
                class = "main-panel",
                style = "padding: 20px;",
                h2("About TagSelect"),
                p("TagSelect is a high-performance application designed to help researchers efficiently select random samples from participant lists and conduct experimental randomization."),
                
                h3("Features"),
                tags$ul(
                  tags$li("Supports multiple file formats (Excel, CSV, SPSS, Stata, Word)"),
                  tags$li("Optimized for large datasets (up to 10,000 participants)"),
                  tags$li("Simple random sampling and advanced experimental randomization"),
                  tags$li("Multiple export options"),
                  tags$li("User-friendly interface with dark/light theme support")
                ),
                
                h3("Technical Details"),
                p("The application uses R's sample.int() function with Fisher-Yates shuffling for simple randomization and the randomizr package for experimental designs."),
                
                h3("Citation"),
                p("If you use TagSelect in your research, please consider citing it:"),
                tags$blockquote(
                  "Ibrahim, M. M. (2025). TagSelect: High-Capacity Random Sampler [Computer software]."
                ),
                
                div(class = "developer-info",
                    h4("Developer Information:"),
                    p("Mudasir Mohammed Ibrahim"),
                    p("Email: mudassiribrahim30@gmail.com"),
                    p("Version: 2.3 | Capacity: 10,000 participants × 1,000 variables")
                )
              )
            )
          ),
          
          div(class = "footer",
              p("TagSelect | Developed by Mudasir Mohammed Ibrahim"),
              p("© 2025 All Rights Reserved")
          )
      )
  )
)

# Server logic remains the same as before
server <- function(input, output, session) {
  data <- reactiveVal(NULL)
  randomized_data <- reactiveVal(NULL)
  exp_data <- reactiveVal(NULL)
  
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
  
  # Handle file upload for simple randomization
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
  
  # Simple randomization sampling
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
  
  # Update variable choices when experimental data changes
  observe({
    req(exp_data())
    df <- exp_data()
    vars <- names(df)[sapply(df, function(x) is.factor(x) | is.character(x) | is.numeric(x))]
    updateSelectInput(session, "group_var", choices = c("None" = "none", vars))
  })
  
  # Handle experimental data upload
  observeEvent(input$exp_file, {
    req(input$exp_file)
    ext <- tools::file_ext(input$exp_file$name)
    
    tryCatch({
      df <- switch(ext,
                   csv = read.csv(input$exp_file$datapath, stringsAsFactors = FALSE),
                   xlsx = read_excel(input$exp_file$datapath),
                   xls = read_excel(input$exp_file$datapath),
                   sav = read_sav(input$exp_file$datapath),
                   dta = read_dta(input$exp_file$datapath),
                   docx = {
                     tmp <- officer::read_docx(input$exp_file$datapath)
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
      
      exp_data(df)
      updateNumericInput(session, "total_participants", value = nrow(df))
      showNotification("Participant list uploaded successfully!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # Create UI for group allocation when using participant list
  output$group_allocation_ui <- renderUI({
    req(exp_data(), input$group_var, input$group_var != "none")
    df <- exp_data()
    groups <- unique(df[[input$group_var]])
    
    div(
      class = "group-allocation-box",
      lapply(groups, function(group) {
        div(
          class = "group-allocation-row",
          numericInput(
            inputId = paste0("n_", gsub("[^[:alnum:]]", "_", group)),
            label = paste("Sample size for", group),
            value = sum(df[[input$group_var]] == group),
            min = 1,
            max = sum(df[[input$group_var]] == group),
            step = 1
          )
        )
      })
    )
  })
  
  # Create UI for group sizes when generating participants
  output$group_size_ui <- renderUI({
    req(input$has_groups, input$group_names)
    groups <- trimws(unlist(strsplit(input$group_names, ",")))
    total_n <- input$total_n
    
    div(
      class = "group-allocation-box",
      lapply(groups, function(group) {
        div(
          class = "group-allocation-row",
          numericInput(
            inputId = paste0("size_", gsub("[^[:alnum:]]", "_", group)),
            label = paste("Number of", group),
            value = ifelse(length(groups) > 0, floor(total_n/length(groups)), total_n),
            min = 1,
            max = total_n,
            step = 1
          )
        )
      })
    )
  })
  
  # Experimental randomization
  observeEvent(input$randomize_btn, {
    tryCatch({
      treatment_names <- trimws(unlist(strsplit(input$treatment_names, ",")))
      if (length(treatment_names) < input$num_treatments) {
        treatment_names <- c(treatment_names, 
                             paste0("Treatment ", (length(treatment_names)+1):input$num_treatments))
      }
      
      if (input$has_participant_list == "yes") {
        req(exp_data())
        df <- exp_data()
        
        # Apply total participants limit if specified
        if (!is.null(input$total_participants)) {
          total_participants <- min(input$total_participants, nrow(df))
          if (total_participants < nrow(df)) {
            df <- df[sample(nrow(df), total_participants), ]
          }
        }
        
        if (input$group_var != "none" && input$specify_group_n && !is.null(input$group_var)) {
          groups <- unique(df[[input$group_var]])
          sampled_rows <- c()
          
          for (group in groups) {
            group_rows <- which(df[[input$group_var]] == group)
            n_input <- input[[paste0("n_", gsub("[^[:alnum:]]", "_", group))]]
            n <- min(n_input, length(group_rows))
            
            if (n > 0) {
              sampled_rows <- c(sampled_rows, sample(group_rows, n))
            }
          }
          
          df <- df[sampled_rows, ]
        }
      } else {
        # Generate synthetic data
        total_n <- input$total_n
        
        if (input$has_groups && !is.null(input$group_names) && input$group_names != "") {
          groups <- trimws(unlist(strsplit(input$group_names, ",")))
          group_sizes <- sapply(groups, function(group) {
            input[[paste0("size_", gsub("[^[:alnum:]]", "_", group))]]
          })
          
          if (sum(group_sizes) > total_n) {
            showNotification("Total group sizes exceed specified total number of participants. Adjusting to fit.", 
                             type = "warning")
            group_sizes <- round(group_sizes * total_n / sum(group_sizes))
          }
          
          df <- data.frame(
            ID = 1:total_n,
            Group = rep(groups, times = group_sizes)[1:total_n]
          )
        } else {
          df <- data.frame(
            ID = 1:total_n
          )
        }
      }
      
      # Perform randomization
      if (input$randomization_method == "simple") {
        assignment <- simple_ra(N = nrow(df), num_arms = input$num_treatments)
      } else if (input$randomization_method == "complete") {
        assignment <- complete_ra(N = nrow(df), num_arms = input$num_treatments)
      } else if (input$randomization_method == "block") {
        if (input$has_participant_list == "yes" && input$group_var != "none") {
          blocks <- df[[input$group_var]]
        } else if (input$has_groups) {
          blocks <- df$Group
        } else {
          blocks <- rep(1:ceiling(nrow(df)/input$block_size), 
                        each = input$block_size)[1:nrow(df)]
        }
        assignment <- block_ra(blocks = blocks, num_arms = input$num_treatments)
      } else if (input$randomization_method == "cluster") {
        if (input$has_participant_list == "yes" && input$group_var != "none") {
          clusters <- df[[input$group_var]]
        } else if (input$has_groups) {
          clusters <- df$Group
        } else {
          showNotification("No cluster variable specified. Using IDs as clusters.", 
                           type = "warning")
          clusters <- df$ID
        }
        assignment <- cluster_ra(clusters = clusters, num_arms = input$num_treatments)
      } else if (input$randomization_method == "stratified") {
        if (input$has_participant_list == "yes" && input$group_var != "none") {
          strata <- df[[input$group_var]]
        } else if (input$has_groups) {
          strata <- df$Group
        } else {
          showNotification("No stratification variable specified. Using simple randomization instead.", 
                           type = "warning")
          strata <- NULL
          assignment <- simple_ra(N = nrow(df), num_arms = input$num_treatments)
        }
        if (!is.null(strata)) {
          assignment <- strata_ra(strata = strata, num_arms = input$num_treatments)
        }
      }
      
      df$Assignment <- treatment_names[assignment]
      randomized_data(df)
      
      # Create summary
      output$randomization_summary <- renderPrint({
        cat("Randomization Method:", input$randomization_method, "\n")
        cat("Number of Treatment Groups:", input$num_treatments, "\n")
        cat("Treatment Group Names:", paste(treatment_names, collapse = ", "), "\n\n")
        
        if (input$has_participant_list == "yes") {
          if (input$group_var != "none" && input$specify_group_n) {
            cat("Group Allocation:\n")
            groups <- unique(df[[input$group_var]])
            for (group in groups) {
              cat(paste0(group, ": ", sum(df[[input$group_var]] == group), "\n"))
            }
            cat("\n")
          }
        } else {
          if (input$has_groups) {
            cat("Group Sizes:\n")
            groups <- trimws(unlist(strsplit(input$group_names, ",")))
            for (group in groups) {
              cat(paste0(group, ": ", sum(df$Group == group), "\n"))
            }
            cat("\n")
          }
        }
        
        cat("Final Assignment Counts:\n")
        print(table(df$Assignment))
      })
      
      showNotification("Randomization completed successfully!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error in randomization:", e$message), type = "error")
    })
  })
  
  output$file_info_box <- renderUI({
    req(input$file)
    
    size <- file.size(input$file$datapath)
    size_text <- format(structure(size, class = "object_size"), units = "auto")
    
    if (size > 5e6) {  # 5MB
      size_text <- tagList(
        span(size_text, style = "color:var(--danger-color);"),
        icon("exclamation-triangle", style = "color:var(--danger-color);")
      )
    }
    
    tagList(
      div(style = "margin-bottom:5px;",
          strong("File: "), input$file$name
      ),
      div(style = "margin-bottom:5px;",
          strong("Size: "), size_text
      ),
      if (!is.null(data())) {
        div(
          strong("Dimensions: "), 
          nrow(data()), "participants ×", ncol(data()), "variables"
        )
      }
    )
  })
  
  output$total_count <- renderText({
    if (is.null(data())) {
      "0"
    } else {
      format(nrow(data()), big.mark = ",")
    }
  })
  
  output$selected_count <- renderText({
    if (is.null(data())) {
      "0"
    } else {
      format(sum(data()$Selected), big.mark = ",")
    }
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
        deferRender = TRUE,
        scroller = TRUE,
        rowCallback = JS(
          "function(row, data) {",
          "  if (data[data.length - 1] === true) {",
          "    $(row).addClass('selected-row');",
          "  }",
          "}")
      ),
      rownames = FALSE,
      filter = 'top',
      class = 'hover stripe nowrap',
      selection = 'none'
    )
  })
  
  output$randomized_table <- renderDT({
    req(randomized_data())
    datatable(
      randomized_data(),
      options = list(
        pageLength = 10,
        scrollX = TRUE,
        scrollY = "500px",
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel', 'pdf'),
        deferRender = TRUE,
        scroller = TRUE
      ),
      rownames = FALSE,
      filter = 'top',
      class = 'hover stripe nowrap',
      selection = 'none'
    )
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
  
  # CSV downloads
  output$download_all_csv <- downloadHandler(
    filename = function() { 
      paste0("FullList_Tagged_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv") 
    },
    content = function(file) {
      write.csv(data(), file, row.names = FALSE)
    }
  )
  
  output$download_selected_csv <- downloadHandler(
    filename = function() { 
      paste0("SelectedParticipants_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv") 
    },
    content = function(file) {
      write.csv(data() %>% filter(Selected == TRUE), file, row.names = FALSE)
    }
  )
  
  # Randomized data download
  output$download_randomized <- downloadHandler(
    filename = function() {
      paste0("RandomizedData_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv")
    },
    content = function(file) {
      write.csv(randomized_data(), file, row.names = FALSE)
    }
  )
}

shinyApp(ui, server)
