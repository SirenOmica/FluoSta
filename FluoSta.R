# FluoSta
# Author: Nokhova A.R.
# If you use FluoSta in your research, please cite: 
# Elfimov K, Gotfrid L, Nokhova A, Gashnikova M, Baboshko D, Totmenin A, Agaphonov A, Gashnikova N. A Semi-Automated Imaging Flow Cytometry Workflow for High-Throughput Quantification of Compound Internalization with IDEAS and FluoSta Software. Methods and Protocols. 2025; 8(6):138. https://doi.org/10.3390/mps8060138

# --- Packages -----------------------------------------------------------------
packages <- c(
  "shiny", "openxlsx", "ez", "tools", "dplyr", "effectsize",
  "plotly", "RColorBrewer", "shinybusy"
)

for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  } else {
    library(pkg, character.only = TRUE)
  }
}

# --- UI -----------------------------------------------------------------------
ui <- fluidPage(
  div(
    style = "background-color: #FFF4E6; padding: 9px 14px; border-radius: 6px; display: inline-block;",
    p(style = "color: #000000; margin: 0;", "If you use FluoSta in your research, please cite: "),
    p(style = "color: #E7670E; font-weight: bold; margin: 0;", "Elfimov K, Gotfrid L, Nokhova A, Gashnikova M, Baboshko D, Totmenin A, Agaphonov A, Gashnikova N. A Semi-Automated Imaging Flow Cytometry Workflow for High-Throughput Quantification of Compound Internalization with IDEAS and FluoSta Software. Methods and Protocols. 2025; 8(6):138. https://doi.org/10.3390/mps8060138")
  ),
  p("For technical support regarding FluoSta, contact: alina.nokhova@gmail.com"),
  
  add_busy_spinner(
    spin = "fading-circle",
    position = "top-right",
    margins = c("0px", "0px"),
    width = "200px",
    height = "200px",
    color = "#FF9500"
  ),
  
  titlePanel("Data analysis from *.txt"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("files", "Select .txt files", multiple = TRUE, accept = ".txt"),
      uiOutput("time_inputs"),
      uiOutput("time_order_select"),
      
      radioButtons("plot_mode", "Plot mode:",
                   choices = c("By substance (one substance at all time)" = "by_substance",
                               "By time (all substances at one time)" = "by_time"),
                   selected = "by_substance"),
      
      actionButton("run_analysis", "Run analysis"),
      br(),
      downloadButton("downloadData", "Download Excel (formatted)"),
      br(), br(),
      downloadButton("downloadRaw", "Download Excel (raw)")
    ),
    
    mainPanel(
      tags$style(HTML(
        "pre#summary_message { background-color: #f8f8f8; padding: 10px; border: 1px solid #ddd; white-space: pre-wrap; font-weight: bold; }"
      )),
      
      verbatimTextOutput("summary_message"),
      
      tabsetPanel(
        tabPanel("Descriptive statistics", tableOutput("desc_stats")),
        tabPanel("ANOVA & Tukey", tableOutput("comp_stats")),
        tabPanel("RM-ANOVA & t-tests", tableOutput("rm_results")),
        tabPanel("Excluded tubes/samples", tableOutput("excluded_samples")),
        tabPanel("Plots", uiOutput("feature_plots"))
      )
    )
  )
)

# --- Server -------------------------------------------------------------------
server <- function(input, output, session) {
  # reactives
  data <- reactiveVal(NULL)
  desc_stats <- reactiveVal(NULL)
  comp_stats <- reactiveVal(NULL)
  rm_results <- reactiveVal(NULL)
  excluded_samples <- reactiveVal(NULL)
  summary_message <- reactiveVal("")
  time_order <- reactiveVal(NULL)
  global_mods <- reactiveVal(NULL)
  
  temp_dir_session <- reactiveVal(NULL)
  results_paths <- reactiveVal(list(formatted = NULL, raw = NULL))
  
  # Helper: try to find app directory
  get_app_dir <- function() {
    args <- commandArgs(trailingOnly = FALSE)
    file_arg <- args[grepl("^--file=", args)]
    if (length(file_arg) > 0) {
      return(normalizePath(dirname(sub("^--file=", "", file_arg[1]))))
    }
    of <- tryCatch(if (!is.null(sys.frames()[[1]]$ofile)) sys.frames()[[1]]$ofile else NULL, error = function(e) NULL)
    if (!is.null(of)) return(normalizePath(dirname(of)))
    return(normalizePath(getwd()))
  }
  
  # Logger writing to analysis_log.txt in app dir
  log_message <- function(msg, type = "INFO") {
    timestamp <- format(Sys.time(), "%Y-%m-%d %H:%M:%S")
    log_entry <- paste0("[", timestamp, "] [", type, "] ", msg, "\n")
    log_file <- file.path(get_app_dir(), "analysis_log.txt")
    tryCatch({
      write(log_entry, file = log_file, append = TRUE)
      TRUE
    }, error = function(e) {
      cat("Error writing to log file:", e$message, "\n")
      FALSE
    })
  }
  
  # --- UI dynamic helpers ----------------------------------------------------
  output$time_inputs <- renderUI({
    req(input$files)
    lapply(seq_along(input$files$name), function(i) {
      textInput(paste0("time_", i), paste0("Time for '", input$files$name[i], "'"), value = "")
    })
  })
  
  output$time_order_select <- renderUI({
    req(input$files)
    times <- unique(sapply(seq_along(input$files$name), function(i) input[[paste0("time_", i)]]))
    times <- times[times != ""]
    if (length(times) == 0) return(NULL)
    named_times <- setNames(times, times)
    selectInput("time_order_select", "Order of time points (delete and enter the times in the correct order if incorrect)",
                choices = named_times, selected = times, multiple = TRUE)
  })
  
  # --- Initialize analysis on Run -------------------------------------------
  observeEvent(input$run_analysis, {
    data(NULL); desc_stats(NULL); comp_stats(NULL); rm_results(NULL); excluded_samples(NULL); summary_message("")
    
    td <- file.path(tempdir(), paste0("FluoSta_temp_", as.integer(Sys.time())))
    unlink(td, recursive = TRUE, force = TRUE)
    dir.create(td, showWarnings = FALSE, recursive = TRUE)
    temp_dir_session(td)
    
    logf <- file.path(get_app_dir(), "analysis_log.txt")
    unlink(logf)
    if (!log_message("Starting new analysis", "INFO")) {
      summary_message(sprintf("\n\nError: could not write to log file\nDetails in '%s'\n\n", logf))
    } else {
      summary_message(sprintf("\n\nAnalysis started\nTemp dir: %s\nDetails in '%s'\n\n", td, logf))
    }
  })
  
  # --- Copy uploaded files and start processing ------------------------------
  observeEvent(input$run_analysis, {
    req(input$files)
    td <- temp_dir_session()
    if (is.null(td) || !dir.exists(td)) {
      summary_message("Internal error: temp dir not initialized. Press Run analysis again.")
      return()
    }
    
    orig <- input$files$name
    times <- sapply(seq_along(orig), function(i) input[[paste0("time_", i)]])
    time_order(unique(times))
    
    if (any(is.na(times) | times == "")) {
      log_message("ERROR: All Time fields must be filled", "ERROR")
      summary_message(sprintf("\n\nError: All Time fields must be filled\nDetails in '%s'\n\n", file.path(get_app_dir(), "analysis_log.txt")))
      return()
    }
    
    for (i in seq_along(orig)) {
      base <- file_path_sans_ext(orig[i])
      newfn <- file.path(td, paste0(base, "_", times[i], "_time.txt"))
      if (file.exists(newfn)) {
        log_message(paste("WARNING: Already exists:", basename(newfn)), "WARNING")
      } else {
        file.copy(input$files$datapath[i], newfn, copy.mode = TRUE)
        log_message(paste("INFO: Created:", newfn), "INFO")
      }
    }
    
    load_and_process_files(td)
  })
  
  # --- File loading and preprocessing ---------------------------------------
  load_and_process_files <- function(temp_dir) {
    files <- list.files(temp_dir, pattern = "_[^_]+_time\\.txt$", full.names = TRUE, ignore.case = TRUE)
    if (!length(files)) {
      log_message("ERROR: No files matching *_<time>_time.txt found", "ERROR")
      summary_message(sprintf("\n\nError: no files for analysis\nDetails in '%s'\n\n", file.path(get_app_dir(), "analysis_log.txt")))
      return()
    }
    
    all <- list()
    for (fn in files) {
      log_message(paste("INFO: Processing:", basename(fn)), "INFO")
      
      t <- sub(".*_([^_]+)_time\\.txt$", "\\1", basename(fn))
      raw <- readLines(fn, warn = FALSE)
      # find header line starting with 'File' (allow leading spaces)
      hdr <- grep("^\\s*File\\b", raw)[1]
      if (is.na(hdr)) {
        log_message(paste("WARNING: file", basename(fn), "does not contain a 'File' header line — skipping"), "WARNING")
        next
      }
      
      cols <- strsplit(raw[hdr], "\t", fixed = TRUE)[[1]]
      df0 <- read.table(fn, skip = hdr, header = FALSE, sep = "\t", dec = ",", fill = TRUE, stringsAsFactors = FALSE, quote = "", comment.char = "")
      if (ncol(df0) < length(cols)) {
        log_message(paste("WARNING: in", basename(fn), "read", ncol(df0), "columns but expected", length(cols), "— skipping"), "WARNING")
        next
      }
      
      df0 <- df0[, seq_len(length(cols)), drop = FALSE]
      colnames(df0) <- cols
      
      if (!"File" %in% cols) {
        log_message(paste("ERROR: file", basename(fn), "is missing 'File' column — skipping"), "ERROR")
        next
      }
      
      valid <- grepl("[-_][0-9]+\\.daf$", df0$File)
      if (!any(valid)) {
        log_message(paste("WARNING: no valid rows in", basename(fn)), "WARNING")
        next
      }
      
      sub <- df0[valid, ]
      
      # Num and File (substance) extraction
      Num <- as.integer(sub(".*[-_]([0-9]+)\\.daf$", "\\1", sub$File))
      File <- sub("_[0-9]+\\.daf$", "", sub$File)
      File <- gsub("-", "_", File, fixed = TRUE)
      
      df <- cbind(data.frame(Num = Num, File = File, time = t), sub[, setdiff(names(sub), "File")])
      log_message(paste("INFO: Loaded", nrow(df), "rows ×", ncol(df), "columns from", basename(fn)), "INFO")
      
      all[[fn]] <- df
    }
    
    if (!length(all)) {
      log_message("ERROR: No files were successfully loaded.", "ERROR")
      summary_message(sprintf("\n\nError: no files loaded\nDetails in '%s'\n\n", file.path(get_app_dir(), "analysis_log.txt")))
      return()
    }
    
    combined <- do.call(rbind, all)
    
    numeric_cols <- setdiff(names(combined), c("File", "time", "Num"))
    for (col in numeric_cols) {
      combined[[col]] <- as.numeric(gsub(",", ".", as.character(combined[[col]])))
    }
    
    data(combined)
    global_mods(sort(unique(combined$File)))
    log_message("INFO: Final dataset is ready", "INFO")
    
    run_analysis(temp_dir)
  }
  
  # --- Small helpers for interpretation ------------------------------------
  add_significance <- function(p) {
    ifelse(is.na(p), "",
           ifelse(p < 0.001, "***",
                  ifelse(p < 0.01, "**",
                         ifelse(p < 0.05, "*", ""))))
  }
  
  add_interpretation <- function(value, type = "omega2", p_value = NA, ci_lower = NA) {
    if (is.na(value) || is.na(p_value) || p_value >= 0.05) return(NA)
    numeric_ci_lower <- suppressWarnings(as.numeric(ci_lower))
    ci_warning <- ifelse(!is.na(numeric_ci_lower) && numeric_ci_lower == 0,
                         " (Caution: Lower confidence interval is 0, indicating unreliability of the effect size estimate.)",
                         "")
    if (type == "omega2") {
      if (value < 0.01) {
        return(paste0("The effect of substances on the variable is very small for this time.", ci_warning))
      } else if (value < 0.06) {
        return(paste0("The effect of substances on the variable is small for this time.", ci_warning))
      } else if (value < 0.14) {
        return(paste0("The effect of substances on the variable is moderate for this time.", ci_warning))
      } else {
        return(paste0("The effect of substances on the variable is large for this time.", ci_warning))
      }
    } else if (type == "ges") {
      if (value < 0.01) {
        return(paste0("The effect of time on the variable is very small for this substance.", ci_warning))
      } else if (value < 0.06) {
        return(paste0("The effect of time on the variable is small for this substance.", ci_warning))
      } else if (value < 0.14) {
        return(paste0("The effect of time on the variable is moderate for this substance.", ci_warning))
      } else {
        return(paste0("The effect of time on the variable is large for this substance.", ci_warning))
      }
    } else if (type == "cohens_d") {
      if (abs(value) < 0.2) {
        return("Very small difference between 2 substances.")
      } else if (abs(value) < 0.5) {
        return("Small difference between 2 substances.")
      } else if (abs(value) < 0.8) {
        return("Moderate difference between 2 substances.")
      } else {
        return("Large difference between 2 substances.")
      }
    } else if (type == "cohens_dz") {
      if (abs(value) < 0.2) {
        return("Very small difference between 2 times.")
      } else if (abs(value) < 0.5) {
        return("Small difference between 2 times.")
      } else if (abs(value) < 0.8) {
        return("Moderate difference between 2 times.")
      } else {
        return("Large difference between 2 times.")
      }
    }
    return(NA)
  }
  
  # --- Main analysis pipeline -----------------------------------------------
  run_analysis <- function(temp_dir) {
    df <- data()
    if (is.null(df)) {
      log_message("ERROR: Please process files first", "ERROR")
      summary_message(sprintf("\n\nError: process files first\nDetails in '%s'\n\n", file.path(get_app_dir(), "analysis_log.txt")))
      return()
    }
    
    levels_time <- time_order()
    if (is.null(levels_time) || length(levels_time) == 0) levels_time <- sort(unique(df$time))
    df$time <- factor(df$time, levels = levels_time)
    df$File <- factor(df$File); df$Num <- factor(df$Num)
    
    feats <- setdiff(names(df), c("Num", "File", "time"))
    if (!length(feats)) {
      log_message("ERROR: No features to analyze!", "ERROR")
      summary_message(sprintf("\n\nError: no features for analysis\nDetails in '%s'\n\n", file.path(get_app_dir(), "analysis_log.txt")))
      return()
    }
    
    desc_df <- data.frame(); comp_df <- data.frame(); rm_df <- data.frame()
    excluded_samples_df <- data.frame()
    
    # Descriptive statistics and ANOVA/Tukey
    for (t in levels(df$time)) {
      sub_t <- subset(df, time == t)
      for (feat in feats) {
        vals <- sub_t[[feat]]
        
        if (all(is.na(vals))) {
          comp_df <- rbind(comp_df, data.frame(Comparison = "Comparison of all substances at this time", Time = t, Feature = feat,
                                               F_from_ANOVA = NA, p_for_ANOVA = NA, Omega2 = NA, Omega2_CI_Lower = "", Omega2_CI_Upper = "", p_Tukey = NA, Cohens_d = NA,
                                               Significance = NA, Comments = NA, Comp_Type = "all", stringsAsFactors = FALSE))
          next
        }
        
        valid_mods <- sapply(levels(sub_t$File), function(mod) {
          mod_vals <- sub_t[[feat]][sub_t$File == mod]
          !all(is.na(mod_vals)) && length(na.omit(mod_vals)) >= 1
        })
        
        sub_t_valid <- sub_t[sub_t$File %in% levels(sub_t$File)[valid_mods], , drop = FALSE]
        
        means <- tapply(vals, sub_t$File, mean, na.rm = TRUE)
        sds <- tapply(vals, sub_t$File, sd, na.rm = TRUE)
        for (g in names(means)) {
          if (!all(is.na(sub_t[[feat]][sub_t$File == g])) && length(na.omit(sub_t[[feat]][sub_t$File == g])) >= 1) {
            desc_df <- rbind(desc_df, data.frame(Substance = g, Time = t, Feature = feat,
                                                 Mean = round(means[g], 4), SD = ifelse(is.na(sds[g]), 0, round(sds[g], 5)), stringsAsFactors = FALSE))
          }
        }
        
        # If there are fewer than 2 distinct valid substances, ANOVA across substances is not applicable
        unique_mods <- unique(sub_t_valid$File)
        unique_mods <- unique_mods[!is.na(unique_mods)]
        if (length(unique_mods) < 2) {
          comp_df <- rbind(comp_df, data.frame(Comparison = "Comparison of all substances at this time", Time = t, Feature = feat,
                                               F_from_ANOVA = NA, p_for_ANOVA = NA, Omega2 = NA, Omega2_CI_Lower = "", Omega2_CI_Upper = "", p_Tukey = NA, Cohens_d = NA,
                                               Significance = "", Comments = "Only one substance present at this time: ANOVA not applicable", Comp_Type = "all", stringsAsFactors = FALSE))
          log_message(paste("INFO:", feat, "at time", t, "- only one substance present, skipping across-substance ANOVA"), "INFO")
          next
        }
        
        if (length(na.omit(vals)) < 2 || var(vals, na.rm = TRUE) == 0) {
          comp_df <- rbind(comp_df, data.frame(Comparison = "Comparison of all substances at this time", Time = t, Feature = feat,
                                               F_from_ANOVA = NA, p_for_ANOVA = NA, Omega2 = NA, Omega2_CI_Lower = "", Omega2_CI_Upper = "", p_Tukey = NA, Cohens_d = NA,
                                               Significance = NA, Comments = "Zero variance or insufficient data: ANOVA not applicable", Comp_Type = "all", stringsAsFactors = FALSE))
          log_message(paste("WARNING:", feat, "at time", t, "- zero variance or insufficient data for ANOVA"), "WARNING")
          next
        }
        
        aov_m <- tryCatch(aov(as.formula(paste0("`", feat, "` ~ `File`")), data = sub_t_valid, na.action = na.omit),
                          error = function(e) { log_message(paste("ERROR: aov error:", e$message), "ERROR"); NULL })
        
        if (!is.null(aov_m)) {
          tab <- summary(aov_m)[[1]]
          Fv <- tab["File", "F value"][1]
          pv <- ifelse(is.na(tab["File", "Pr(>F)" ]), NA, round(tab["File", "Pr(>F)" ], 4)[1])
          
          ssb <- tab["File", "Sum Sq"][1]; dfb <- tab["File", "Df"][1]; msw <- tab["Residuals", "Mean Sq"][1]; ssw <- tab["Residuals", "Sum Sq"][1]
          om2 <- ifelse(is.na(ssb) || is.na(dfb) || is.na(msw) || is.na(ssw), NA, round((ssb - dfb * msw)/(ssb + ssw + msw), 3))
          
          ci_om2 <- tryCatch({ omega_ci <- omega_squared(aov_m, ci = 0.95); if (!is.null(omega_ci)) c(omega_ci$CI_low, omega_ci$CI_high) else c(NA, NA) },
                             error = function(e) c(NA, NA))
          if (is.na(ci_om2[1]) || is.na(ci_om2[2])) ci_om2 <- c("", "") else ci_om2 <- round(ci_om2, 3)
          
          comp_df <- rbind(comp_df, data.frame(Comparison = "Comparison of all substances at this time", Time = t, Feature = feat,
                                               F_from_ANOVA = ifelse(is.na(Fv), NA, round(Fv, 2)), p_for_ANOVA = pv, Omega2 = om2,
                                               Omega2_CI_Lower = ci_om2[1], Omega2_CI_Upper = ci_om2[2], p_Tukey = NA, Cohens_d = NA,
                                               Significance = add_significance(pv),
                                               Comments = add_interpretation(om2, "omega2", pv, ci_om2[1]),
                                               Comp_Type = "all", stringsAsFactors = FALSE))
          
          tuk <- tryCatch({ TukeyHSD(aov_m, "File")$File }, error = function(e) { log_message(paste("WARNING: TukeyHSD failed for", feat, "at time", t, "->", e$message), "WARNING"); NULL })
          if (!is.null(tuk)) {
            for (row in rownames(tuk)) {
              parts <- strsplit(row, "-", fixed = TRUE)[[1]]
              cohens_d_val <- NA
              try({
                group1 <- sub_t_valid[sub_t_valid$File == parts[2], feat]
                group2 <- sub_t_valid[sub_t_valid$File == parts[1], feat]
                if (length(na.omit(group1)) >= 2 && length(na.omit(group2)) >= 2) {
                  cohens_d_val <- cohens_d(group1, group2)$Cohens_d
                  cohens_d_val <- round(cohens_d_val, 3)
                }
              }, silent = TRUE)
              
              comp_df <- rbind(comp_df, data.frame(Comparison = paste(parts[2], "vs", parts[1]),
                                                   Time = t, Feature = feat, F_from_ANOVA = NA, p_for_ANOVA = NA,
                                                   Omega2 = NA, Omega2_CI_Lower = "", Omega2_CI_Upper = "",
                                                   p_Tukey = ifelse(is.na(tuk[row, "p adj"]), NA, round(tuk[row, "p adj"], 4)),
                                                   Cohens_d = cohens_d_val,
                                                   Significance = add_significance(tuk[row, "p adj"]),
                                                   Comments = add_interpretation(cohens_d_val, "cohens_d", tuk[row, "p adj"]),
                                                   Comp_Type = "pairwise", stringsAsFactors = FALSE))
            }
          }
        }
      }
    }
    
    # Repeated-measures ANOVA & paired t-tests
    for (mod in levels(df$File)) {
      sub_mod <- subset(df, File == mod)
      time_levels <- levels(df$time)
      
      if (nrow(sub_mod) == 0) {
        for (feat in feats) {
          rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = "Comparison of all times for this substance", Feature = feat,
                                           F_from_RM_ANOVA = NA, p_for_RM_ANOVA = NA, GES = NA, GES_CI_Lower = "", GES_CI_Upper = "",
                                           p_pairwise_t_test = NA, Cohens_dz = NA, Significance = NA, Comments = "No data for Substance", Comp_Type = "all", stringsAsFactors = FALSE))
          excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = "All", Feature = feat, Reason = "No data", stringsAsFactors = FALSE))
        }
        log_message(paste("WARNING:", mod, "- no data for this Substance"), "WARNING")
        next
      }
      
      nums_valid <- sub_mod %>% group_by(Num) %>% summarize(n_times = n_distinct(time)) %>% filter(n_times >= 2) %>% pull(Num)
      excluded_nums <- setdiff(unique(sub_mod$Num), nums_valid)
      if (length(excluded_nums) > 0) {
        for (feat in feats) {
          excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(excluded_nums, collapse = ", "), Feature = feat, Reason = "Missing in some time points", stringsAsFactors = FALSE))
        }
        log_message(paste("WARNING: For Substance", mod, "excluded Num(s):", paste(excluded_nums, collapse = ", "), "- missing in some time points"), "WARNING")
      }
      
      sub_mod_complete <- sub_mod[sub_mod$Num %in% nums_valid, ]
      
      if (length(nums_valid) < 2) {
        for (feat in feats) {
          rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = "Comparison of all times for this substance", Feature = feat,
                                           F_from_RM_ANOVA = NA, p_for_RM_ANOVA = NA, GES = NA, GES_CI_Lower = "", GES_CI_Upper = "",
                                           p_pairwise_t_test = NA, Cohens_dz = NA, Significance = NA, Comments = "Insufficient samples (Num < 2)", Comp_Type = "all", stringsAsFactors = FALSE))
          excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(nums_valid, collapse = ", "), Feature = feat, Reason = "Insufficient number of samples", stringsAsFactors = FALSE))
        }
        log_message(paste("WARNING: For Substance", mod, "- insufficient number of samples (Num < 2)"), "WARNING")
        next
      }
      
      for (feat in feats) {
        sub <- sub_mod_complete[, c("Num", "time", feat)]
        names(sub)[3] <- "value"
        
        if (nrow(sub) == 0 || all(is.na(sub$value))) {
          rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = "Comparison of all times for this substance", Feature = feat,
                                           F_from_RM_ANOVA = NA, p_for_RM_ANOVA = NA, GES = NA, GES_CI_Lower = "", GES_CI_Upper = "",
                                           p_pairwise_t_test = NA, Cohens_dz = NA, Significance = NA, Comments = "No valid data", Comp_Type = "all", stringsAsFactors = FALSE))
          excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(nums_valid, collapse = ", "), Feature = feat, Reason = "No valid data or all NA", stringsAsFactors = FALSE))
          log_message(paste("WARNING:", mod, feat, "- no valid data or all NA for RM-ANOVA"), "WARNING")
          next
        }
        
        valid_values <- sub$value[!is.na(sub$value)]
        if (length(valid_values) < 2) {
          rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = "Comparison of all times for this substance", Feature = feat,
                                           F_from_RM_ANOVA = NA, p_for_RM_ANOVA = NA, GES = NA, GES_CI_Lower = "", GES_CI_Upper = "",
                                           p_pairwise_t_test = NA, Cohens_dz = NA, Significance = NA, Comments = "Insufficient valid data", Comp_Type = "all", stringsAsFactors = FALSE))
          excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(nums_valid, collapse = ", "), Feature = feat, Reason = "Insufficient valid data", stringsAsFactors = FALSE))
          log_message(paste("WARNING:", mod, feat, "- insufficient valid data for RM-ANOVA"), "WARNING")
          next
        }
        
        if (var(valid_values, na.rm = TRUE) == 0) {
          rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = "Comparison of all times for this substance", Feature = feat,
                                           F_from_RM_ANOVA = NA, p_for_RM_ANOVA = NA, GES = NA, GES_CI_Lower = "", GES_CI_Upper = "",
                                           p_pairwise_t_test = NA, Cohens_dz = NA, Significance = NA, Comments = "Zero variance", Comp_Type = "all", stringsAsFactors = FALSE))
          excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(nums_valid, collapse = ", "), Feature = feat, Reason = "Zero variance", stringsAsFactors = FALSE))
          log_message(paste("WARNING:", mod, feat, "- zero variance for RM-ANOVA"), "WARNING")
          next
        }
        
        time_counts <- sub %>% group_by(time) %>% summarize(n = sum(!is.na(value))) %>% filter(n >= 2)
        if (nrow(time_counts) < 2) {
          rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = "Comparison of all times for this substance", Feature = feat,
                                           F_from_RM_ANOVA = NA, p_for_RM_ANOVA = NA, GES = NA, GES_CI_Lower = "", GES_CI_Upper = "",
                                           p_pairwise_t_test = NA, Cohens_dz = NA, Significance = NA, Comments = "Insufficient time points with data", Comp_Type = "all", stringsAsFactors = FALSE))
          excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(nums_valid, collapse = ", "), Feature = feat, Reason = "Insufficient time points with data", stringsAsFactors = FALSE))
          log_message(paste("WARNING:", mod, feat, "- insufficient time points with data for RM-ANOVA"), "WARNING")
          next
        }
        
        aov_m <- tryCatch(aov(value ~ time + Error(Num/time), data = sub, na.action = na.omit),
                          error = function(e) { log_message(paste("ERROR: aov rm error:", e$message), "ERROR"); NULL })
        if (is.null(aov_m)) next
        
        tab <- summary(aov_m)$`Error: Num:time`[[1]]
        Fv <- tab["time", "F value"]
        pv <- ifelse(is.na(tab["time", "Pr(>F)"]), NA, round(tab["time", "Pr(>F)"], 4))
        ss_time <- tab["time", "Sum Sq"]; ss_error <- tab["Residuals", "Sum Sq"]
        ges <- ifelse(is.na(ss_time) || is.na(ss_error), NA, round(ss_time / (ss_time + ss_error), 3))
        
        ci_ges <- tryCatch({ ges_ci <- eta_squared(aov_m, generalized = TRUE, ci = 0.95); if (!is.null(ges_ci)) c(ges_ci$CI_low, ges_ci$CI_high) else c(NA, NA)},
                           error = function(e) c(NA, NA))
        if (is.na(ci_ges[1]) || is.na(ci_ges[2])) ci_ges <- c("", "") else ci_ges <- round(ci_ges, 3)
        
        rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = "Comparison of all times for this substance", Feature = feat,
                                         F_from_RM_ANOVA = ifelse(is.na(Fv), NA, round(Fv, 2)), p_for_RM_ANOVA = pv, GES = ges,
                                         GES_CI_Lower = ci_ges[1], GES_CI_Upper = ci_ges[2], p_pairwise_t_test = NA, Cohens_dz = NA,
                                         Significance = add_significance(pv),
                                         Comments = add_interpretation(ges, "ges", pv, ci_ges[1]), Comp_Type = "all", stringsAsFactors = FALSE))
        
        # paired t-tests
        p_values <- c(); comparisons <- c(); cohens_d_values <- c()
        if (length(time_levels) >= 2) {
          for (i in 1:(length(time_levels) - 1)) for (j in (i + 1):length(time_levels)) {
            t1 <- time_levels[i]; t2 <- time_levels[j]
            sub_t1 <- subset(sub, time == t1); sub_t2 <- subset(sub, time == t2)
            common_nums <- intersect(sub_t1$Num, sub_t2$Num)
            if (length(common_nums) >= 2) {
              vals_t1 <- sub_t1$value[match(common_nums, sub_t1$Num)]
              vals_t2 <- sub_t2$value[match(common_nums, sub_t2$Num)]
              
              if (length(na.omit(vals_t1)) < 2 || length(na.omit(vals_t2)) < 2 || var(vals_t1, na.rm = TRUE) == 0 || var(vals_t2, na.rm = TRUE) == 0) {
                excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(common_nums, collapse = ", "), Feature = feat, Reason = paste("Insufficient data or zero variance for t-test between", t1, "and", t2), stringsAsFactors = FALSE))
                log_message(paste("WARNING:", mod, feat, "- cannot run paired t-test between", t1, "and", t2, "- insufficient data or zero variance"), "WARNING")
                next
              }
              
              t_result <- tryCatch(t.test(vals_t1, vals_t2, paired = TRUE), error = function(e) { excluded_samples_df <<- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(common_nums, collapse = ", "), Feature = feat, Reason = paste("Error in t-test between", t1, "and", t2), stringsAsFactors = FALSE)); log_message(paste("WARNING: t-test error for", mod, feat, "between", t1, "and", t2, "->", e$message), "WARNING"); NULL })
              if (!is.null(t_result)) {
                cohens_val <- tryCatch({ cohens_d(vals_t1, vals_t2, paired = TRUE)$Cohens_d }, error = function(e) { NA })
                p_values <- c(p_values, t_result$p.value)
                comparisons <- c(comparisons, paste(t1, "vs", t2))
                cohens_d_values <- c(cohens_d_values, cohens_val)
              }
            } else {
              excluded_samples_df <- rbind(excluded_samples_df, data.frame(Substance = mod, Excluded_Num = paste(common_nums, collapse = ", "), Feature = feat, Reason = paste("Insufficient common subjects for t-test between", t1, "and", t2), stringsAsFactors = FALSE))
              log_message(paste("WARNING:", mod, feat, "- insufficient common subjects for paired t-test between", t1, "and", t2), "WARNING")
            }
          }
          
          if (length(p_values) > 0) {
            p_adjusted <- round(p.adjust(p_values, method = "BH"), 4)
            for (k in seq_along(p_values)) {
              rm_df <- rbind(rm_df, data.frame(Substance = mod, Time = comparisons[k], Feature = feat,
                                               F_from_RM_ANOVA = NA, p_for_RM_ANOVA = NA, GES = NA, GES_CI_Lower = "", GES_CI_Upper = "",
                                               p_pairwise_t_test = p_adjusted[k], Cohens_dz = cohens_d_values[k],
                                               Significance = add_significance(p_adjusted[k]),
                                               Comments = add_interpretation(cohens_d_values[k], "cohens_dz", p_adjusted[k]), Comp_Type = "pairwise", stringsAsFactors = FALSE))
            }
          }
        }
      }
    }
    
    # Save raw snapshots
    raw_desc_df <- desc_df
    raw_comp_df <- comp_df
    raw_rm_df <- rm_df
    
    # --- Formatting for UI --------------------------------------------------
    ensure_cols <- function(df, cols) {
      missing <- setdiff(cols, names(df))
      if (length(missing) > 0) for (m in missing) df[[m]] <- NA
      df
    }
    
    desc_df <- ensure_cols(desc_df, c("Substance", "Time", "Feature", "Mean", "SD"))
    comp_df <- ensure_cols(comp_df, c("Comparison", "Time", "Feature", "F_from_ANOVA", "p_for_ANOVA", "Omega2", "Omega2_CI_Lower", "Omega2_CI_Upper", "p_Tukey", "Cohens_d", "Significance", "Comments", "Comp_Type"))
    rm_df   <- ensure_cols(rm_df,   c("Substance", "Time", "Feature", "F_from_RM_ANOVA", "p_for_RM_ANOVA", "GES", "GES_CI_Lower", "GES_CI_Upper", "p_pairwise_t_test", "Cohens_dz", "Significance", "Comments", "Comp_Type"))
    
    desc_df <- desc_df %>% mutate(Mean = formatC(as.numeric(Mean), digits = 4, format = "f"), SD = formatC(as.numeric(SD), digits = 5, format = "f"))
    
    safe_format <- function(x, digits = 3, fmt = "f") {
      sapply(x, function(v) {
        nv <- suppressWarnings(as.numeric(v))
        if (is.na(nv)) return("")
        formatC(nv, digits = digits, format = fmt, flag = "")
      }, USE.NAMES = FALSE)
    }
    
    if (nrow(comp_df) > 0) {
      comp_df$F_from_ANOVA   <- ifelse(comp_df$Comp_Type == "pairwise", "", safe_format(comp_df$F_from_ANOVA, 2, "f"))
      comp_df$p_for_ANOVA    <- ifelse(comp_df$Comp_Type == "pairwise", "", safe_format(comp_df$p_for_ANOVA, 4, "f"))
      comp_df$Omega2         <- ifelse(comp_df$Comp_Type == "pairwise", "", safe_format(comp_df$Omega2, 3, "f"))
      comp_df$Omega2_CI_Lower<- ifelse(comp_df$Comp_Type == "pairwise", "", safe_format(comp_df$Omega2_CI_Lower, 3, "f"))
      comp_df$Omega2_CI_Upper<- ifelse(comp_df$Comp_Type == "pairwise", "", safe_format(comp_df$Omega2_CI_Upper, 3, "f"))
      comp_df$p_Tukey        <- ifelse(comp_df$Comp_Type == "all", "", safe_format(comp_df$p_Tukey, 4, "f"))
      comp_df$Cohens_d       <- ifelse(comp_df$Comp_Type == "all", "", safe_format(comp_df$Cohens_d, 3, "f"))
      comp_df$Significance   <- ifelse(is.na(comp_df$Significance), "", comp_df$Significance)
      comp_df$Comments       <- ifelse(is.na(comp_df$Comments), "", comp_df$Comments)
    }
    
    if (nrow(rm_df) > 0) {
      rm_df$F_from_RM_ANOVA  <- ifelse(rm_df$Comp_Type == "pairwise", "", safe_format(rm_df$F_from_RM_ANOVA, 2, "f"))
      rm_df$p_for_RM_ANOVA   <- ifelse(rm_df$Comp_Type == "pairwise", "", safe_format(rm_df$p_for_RM_ANOVA, 4, "f"))
      rm_df$GES              <- ifelse(rm_df$Comp_Type == "pairwise", "", safe_format(rm_df$GES, 3, "f"))
      rm_df$GES_CI_Lower     <- ifelse(rm_df$Comp_Type == "pairwise", "", safe_format(rm_df$GES_CI_Lower, 3, "f"))
      rm_df$GES_CI_Upper     <- ifelse(rm_df$Comp_Type == "pairwise", "", safe_format(rm_df$GES_CI_Upper, 3, "f"))
      rm_df$p_pairwise_t_test<- ifelse(rm_df$Comp_Type == "all", "", safe_format(rm_df$p_pairwise_t_test, 4, "f"))
      rm_df$Cohens_dz        <- ifelse(rm_df$Comp_Type == "all", "", safe_format(rm_df$Cohens_dz, 3, "f"))
      rm_df$Significance     <- ifelse(is.na(rm_df$Significance), "", rm_df$Significance)
      rm_df$Comments         <- ifelse(is.na(rm_df$Comments), "", rm_df$Comments)
    }
    
    desc_stats(desc_df); comp_stats(comp_df); rm_results(rm_df)
    excluded_samples(excluded_samples_df)
    
    output$desc_stats <- renderTable({ desc_df }, digits = 4)
    output$comp_stats <- renderTable({ comp_df }, digits = 4)
    output$rm_results  <- renderTable({ rm_df }, digits = 4)
    output$excluded_samples <- renderTable({
      ex <- excluded_samples_df
      if (is.null(ex) || nrow(ex) == 0) {
        data.frame(Info = "No excluded tubes/samples for RM-ANOVA/t-test")
      } else {
        ex
      }
    })
    
    # --- Plot generation ----------------------------------------------------
    output$feature_plots <- renderUI({
      req(desc_stats())
      desc_df_local <- desc_stats()
      feats <- unique(desc_df_local$Feature)
      tagList(lapply(seq_along(feats), function(i) {
        feat <- feats[i]
        safe_id <- paste0("plot_time_", i, "_", gsub("[^A-Za-z0-9]", "_", feat))
        tagList(
          h3(style = "font-size:20px; font-weight:bold;", paste("Feature:", feat)),
          div(style = "overflow-x:auto;", plotlyOutput(safe_id, height = "600px"))
        )
      }))
    })
    
    get_colors <- function(n) {
      if (n <= 9) RColorBrewer::brewer.pal(n, "Set1") else grDevices::colorRampPalette(RColorBrewer::brewer.pal(9, "Set1"))(n)
    }
    
    hex_to_rgba <- function(hex, alpha = 1) {
      rgb <- grDevices::col2rgb(hex) / 255
      sprintf("rgba(%d,%d,%d,%.3f)", round(rgb[1] * 255), round(rgb[2] * 255), round(rgb[3] * 255), alpha)
    }
    
    wrap_title_for_plot <- function(s, max_len = 100) {
      if (is.na(s) || s == "") return(s)
      if (nchar(s) <= max_len) return(s)
      words <- strsplit(s, "[ .-]")[[1]]
      lines <- c()
      cur <- ""
      for (w in words) {
        if (nchar(cur) + nchar(w) + 1 <= max_len) {
          cur <- if (cur == "") w else paste(cur, w, sep = " ")
        } else {
          lines <- c(lines, cur)
          cur <- w
        }
      }
      if (cur != "") lines <- c(lines, cur)
      paste(lines, collapse = "<br>")
    }
    
    make_plotly_for_feat <- function(feat_local, plot_mode_local, desc_df_local, time_levels_local, mods_local,
                                     marker_size_sub = 14, marker_size_time = 12, alpha_marker = 0.85, alpha_err = 0.6,
                                     divider_alpha = 0.85) {
      sub_df <- desc_df_local %>% filter(Feature == feat_local)
      sub_df$Mean_num <- as.numeric(gsub(",", ".", sub_df$Mean))
      sub_df$SD_num <- as.numeric(gsub(",", ".", sub_df$SD))
      
      cols <- get_colors(length(mods_local))
      names(cols) <- mods_local
      rgba_markers <- sapply(cols, function(h) hex_to_rgba(h, alpha_marker))
      rgba_errs <- sapply(cols, function(h) hex_to_rgba(h, alpha_err))
      
      raw_title <- feat_local
      wrapped_title <- wrap_title_for_plot(raw_title, max_len = 100)
      title_font_size_local <- if (nchar(raw_title) > 100) 16 else 18
      title_html <- paste0("<b>", wrapped_title, "</b>")
      
      if (plot_mode_local == "by_substance") {
        rows <- data.frame()
        idx <- 1
        for (m in mods_local) {
          for (t in time_levels_local) {
            r <- sub_df %>% filter(Substance == m, Time == t)
            if (nrow(r) == 0) {
              rows <- rbind(rows, data.frame(x = idx, Substance = m, time = t, Mean = NA, SD = NA, stringsAsFactors = FALSE))
            } else {
              rows <- rbind(rows, data.frame(x = idx, Substance = m, time = t, Mean = r$Mean_num[1], SD = r$SD_num[1], stringsAsFactors = FALSE))
            }
            idx <- idx + 1
          }
          idx <- idx + 0.5
        }
        
        p <- plot_ly()
        for (i in seq_along(mods_local)) {
          m <- mods_local[i]
          sel <- rows %>% filter(Substance == m)
          p <- add_trace(p,
                         x = sel$x, y = sel$Mean,
                         type = 'scatter', mode = 'lines+markers',
                         name = m,
                         marker = list(size = marker_size_sub, color = rgba_markers[i]),
                         line = list(width = 2, color = cols[i]),
                         error_y = list(type = "data", array = sel$SD, visible = TRUE, width = 2, color = rgba_errs[i])
          )
        }
        
        p <- layout(p,
                    title = list(text = title_html, font = list(size = title_font_size_local)),
                    xaxis = list(title = "", tickvals = rows$x, ticktext = rows$time,
                                 tickangle = 45, tickfont = list(size = 16, color = "black", family = "Arial Black"),
                                 titlefont = list(size = 16), showgrid = FALSE, showline = TRUE, linecolor = "black", linewidth = 2, mirror = "ticks", zeroline = FALSE, zerolinewidth = 0),
                    yaxis = list(title = "", tickfont = list(size = 16, color = "black", family = "Arial Black"),
                                 showline = TRUE, linecolor = "black", linewidth = 2, mirror = "ticks", zeroline = FALSE, zerolinewidth = 0),
                    legend = list(font = list(family = "Arial Black", size = 16)),
                    margin = list(t = 120, b = 140))
        return(p)
        
      } else {
        # by_time
        rows <- data.frame()
        pos <- 1
        group_centers <- c()
        group_ends <- c()
        for (t in time_levels_local) {
          start_pos <- pos
          for (m in mods_local) {
            r <- sub_df %>% filter(Substance == m, Time == t)
            if (nrow(r) == 0) {
              rows <- rbind(rows, data.frame(x = pos, Substance = m, time = t, Mean = NA, SD = NA, stringsAsFactors = FALSE))
            } else {
              rows <- rbind(rows, data.frame(x = pos, Substance = m, time = t, Mean = r$Mean_num[1], SD = r$SD_num[1], stringsAsFactors = FALSE))
            }
            pos <- pos + 1
          }
          end_pos <- pos - 1
          center <- (start_pos + end_pos) / 2
          group_centers <- c(group_centers, center)
          group_ends <- c(group_ends, end_pos)
          pos <- pos + 1.5
        }
        
        gap_positions <- if (length(group_ends) >= 2) sapply(group_ends[-length(group_ends)], function(e) e + 0.75) else numeric(0)
        shapes_list <- lapply(gap_positions, function(gx) {
          list(type = "line",
               x0 = gx, x1 = gx,
               xref = "x",
               y0 = 0, y1 = 1,
               yref = "paper",
               line = list(color = sprintf("rgba(110,110,110,%.2f)", 0.9), width = 1, dash = "dot"))
        })
        
        p <- plot_ly()
        for (i in seq_along(mods_local)) {
          m <- mods_local[i]
          sel <- rows %>% filter(Substance == m)
          p <- add_trace(p,
                         x = sel$x, y = sel$Mean,
                         type = 'scatter', mode = 'markers',
                         name = m,
                         marker = list(size = marker_size_time, color = rgba_markers[i]),
                         error_y = list(type = "data", array = sel$SD, visible = TRUE, width = 2, color = rgba_errs[i])
          )
        }
        
        any_long <- any(nchar(time_levels_local) > 7)
        tick_ang <- ifelse(any_long, 45, 0)
        
        p <- layout(p,
                    title = list(text = title_html, font = list(size = title_font_size_local)),
                    xaxis = list(title = "", tickvals = group_centers, ticktext = time_levels_local,
                                 tickangle = tick_ang, tickfont = list(size = 16, color = "black", family = "Arial Black"),
                                 titlefont = list(size = 16), showgrid = FALSE, showline = TRUE, linecolor = "black", linewidth = 2, mirror = "ticks", zeroline = FALSE, zerolinewidth = 0),
                    yaxis = list(title = "", tickfont = list(size = 16, color = "black", family = "Arial Black"),
                                 showline = TRUE, linecolor = "black", linewidth = 2, mirror = "ticks", zeroline = FALSE, zerolinewidth = 0),
                    legend = list(font = list(family = "Arial Black", size = 16)),
                    shapes = shapes_list,
                    margin = list(t = 120, b = 140))
        return(p)
      }
    }
    
    plots_list <- reactiveVal(list())
    
    observe({
      req(desc_stats(), input$plot_mode, !is.null(input$time_order_select))
      desc_df_local <- desc_stats()
      feats <- unique(desc_df_local$Feature)
      time_levels_local <- input$time_order_select
      if (is.null(time_levels_local) || length(time_levels_local) == 0) time_levels_local <- time_order()
      mods <- global_mods()
      
      marker_size_sub <- 14
      marker_size_time <- 12
      
      plst <- list()
      for (i in seq_along(feats)) {
        feat <- feats[i]
        safe_id <- paste0("plot_time_", i, "_", gsub("[^A-Za-z0-9]", "_", feat))
        
        p_obj <- make_plotly_for_feat(feat, input$plot_mode, desc_df_local, time_levels_local, mods,
                                      marker_size_sub = marker_size_sub,
                                      marker_size_time = marker_size_time,
                                      alpha_marker = 0.85, alpha_err = 0.6,
                                      divider_alpha = 0.85)
        
        plst[[safe_id]] <- p_obj
        
        local({
          sid <- safe_id
          f_local <- feat
          output[[sid]] <- renderPlotly({
            make_plotly_for_feat(f_local, input$plot_mode, desc_df_local, time_levels_local, mods,
                                 marker_size_sub = marker_size_sub,
                                 marker_size_time = marker_size_time,
                                 alpha_marker = 0.85, alpha_err = 0.6,
                                 divider_alpha = 0.85)
          })
        })
      }
      plots_list(plst)
    })
    
    # --- Export Excel -------------------------------------------------------
    wb <- createWorkbook()
    addWorksheet(wb, "descriptive_statistics")
    addWorksheet(wb, "ANOVA_Tukey")
    addWorksheet(wb, "RM-ANOVA_t")
    writeData(wb, "descriptive_statistics", desc_df)
    writeData(wb, "ANOVA_Tukey", comp_df)
    writeData(wb, "RM-ANOVA_t", rm_df)
    
    num_style_2 <- createStyle(numFmt = "0.00")
    num_style_3 <- createStyle(numFmt = "0.000")
    num_style_4 <- createStyle(numFmt = "0.0000")
    num_style_5 <- createStyle(numFmt = "0.00000")
    
    if (nrow(desc_df) > 0) {
      addStyle(wb, "descriptive_statistics", num_style_4, rows = 2:(nrow(desc_df) + 1), cols = 4, gridExpand = TRUE)
      addStyle(wb, "descriptive_statistics", num_style_5, rows = 2:(nrow(desc_df) + 1), cols = 5, gridExpand = TRUE)
    }
    if (nrow(comp_df) > 0) {
      addStyle(wb, "ANOVA_Tukey", num_style_2, rows = 2:(nrow(comp_df) + 1), cols = 4, gridExpand = TRUE)
      addStyle(wb, "ANOVA_Tukey", num_style_3, rows = 2:(nrow(comp_df) + 1), cols = c(6, 7, 8, 10), gridExpand = TRUE)
      addStyle(wb, "ANOVA_Tukey", num_style_4, rows = 2:(nrow(comp_df) + 1), cols = c(5, 9), gridExpand = TRUE)
    }
    if (nrow(rm_df) > 0) {
      addStyle(wb, "RM-ANOVA_t", num_style_2, rows = 2:(nrow(rm_df) + 1), cols = 4, gridExpand = TRUE)
      addStyle(wb, "RM-ANOVA_t", num_style_3, rows = 2:(nrow(rm_df) + 1), cols = c(6, 7, 8, 10), gridExpand = TRUE)
      addStyle(wb, "RM-ANOVA_t", num_style_4, rows = 2:(nrow(rm_df) + 1), cols = c(5, 9), gridExpand = TRUE)
    }
    
    formatted_fp <- file.path(temp_dir, "Results.xlsx")
    raw_fp       <- file.path(temp_dir, "Results_raw.xlsx")
    
    saveWorkbook(wb, formatted_fp, overwrite = TRUE)
    
    wb_raw <- createWorkbook()
    addWorksheet(wb_raw, "descriptive_statistics_raw")
    addWorksheet(wb_raw, "ANOVA_Tukey_raw")
    addWorksheet(wb_raw, "RM-ANOVA_t_raw")
    writeData(wb_raw, "descriptive_statistics_raw", raw_desc_df)
    writeData(wb_raw, "ANOVA_Tukey_raw", raw_comp_df)
    writeData(wb_raw, "RM-ANOVA_t_raw", raw_rm_df)
    saveWorkbook(wb_raw, raw_fp, overwrite = TRUE)
    
    results_paths(list(formatted = formatted_fp, raw = raw_fp))
    
    log_message("INFO: Analysis completed", "INFO")
    summary_message(sprintf("\n\nAnalysis completed\nTemp dir: %s\nDetails in '%s'\n\n", temp_dir, file.path(get_app_dir(), "analysis_log.txt")))
  }
  
  # render summary message
  output$summary_message <- renderText({ summary_message() })
  
  # --- Downloads ------------------------------------------------------------
  output$downloadData <- downloadHandler(
    filename = function() "Results.xlsx",
    content = function(file) {
      rp <- results_paths()
      if (is.null(rp$formatted) || !file.exists(rp$formatted)) {
        stop("Results.xlsx not found. Run analysis first.")
      }
      file.copy(rp$formatted, file, overwrite = TRUE)
    }
  )
  
  output$downloadRaw <- downloadHandler(
    filename = function() "Results_raw.xlsx",
    content = function(file) {
      rp <- results_paths()
      if (is.null(rp$raw) || !file.exists(rp$raw)) {
        stop("Results_raw.xlsx not found. Run analysis first.")
      }
      file.copy(rp$raw, file, overwrite = TRUE)
    }
  )
}

# Launch the application
shinyApp(ui, server)
