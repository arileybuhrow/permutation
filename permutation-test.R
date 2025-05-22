# Set seed for reproducibility
set.seed(42)

# Load required package
library(dplyr)

# Number of samples per group
n_samples <- 30

# Function to generate zero-inflated data
generate_zero_inflated_data <- function(n, mean_val, prob_zero = 0.2, 
                                        sd_val = 0.5) {
  is_nonzero <- rbinom(n, 1, 1 - prob_zero)
  values <- is_nonzero * rnorm(n, mean = mean_val, sd = sd_val)
  return(values)
}

# Generate data for groups A, B, and C
# Groups A and B: same distribution (mean = 4)
group_A <- generate_zero_inflated_data(n_samples, mean_val = 4)
group_B <- generate_zero_inflated_data(n_samples, mean_val = 4)
# Group C: different distribution (mean = 6)
group_C <- generate_zero_inflated_data(n_samples, mean_val = 6)

# Combine data into a single data frame
data <- data.frame(
  Group = rep(c("A", "B", "C"), each = n_samples),
  LogCFU = c(group_A, group_B, group_C)
)

# Perform a permutation test between two groups using the difference in medians
permutation_test <- function(data, group1, group2, nperm = 10000, stat = "median") {
  subset_data <- data[data$Group %in% c(group1, group2), ]
  
  # Calculate observed difference (group2 - group1)
  if(stat == "median") {
    obs_diff <- median(subset_data$LogCFU[subset_data$Group == group2]) -
      median(subset_data$LogCFU[subset_data$Group == group1])
  } else if(stat == "mean") {
    obs_diff <- mean(subset_data$LogCFU[subset_data$Group == group2]) -
      mean(subset_data$LogCFU[subset_data$Group == group1])
  } else {
    stop("Unsupported statistic")
  }
  
  perm_diff <- numeric(nperm)
  for(i in 1:nperm) {
    permuted_labels <- sample(subset_data$Group)
    if(stat == "median") {
      perm_diff[i] <- median(subset_data$LogCFU[permuted_labels == group2]) -
        median(subset_data$LogCFU[permuted_labels == group1])
    } else if(stat == "mean") {
      perm_diff[i] <- mean(subset_data$LogCFU[permuted_labels == group2]) -
        mean(subset_data$LogCFU[permuted_labels == group1])
    }
  }
  
  # Two-sided p-value
  p_value <- mean(abs(perm_diff) >= abs(obs_diff))
  
  return(list(obs_diff = obs_diff, p_value = p_value, perm_diff = perm_diff))
}

# Define pairwise comparisons: A vs B (no true difference), A vs C and B vs C (true differences)
group_pairs <- list(c("A", "B"), c("A", "C"), c("B", "C"))
results <- list()

nperm <- 10000

for(pair in group_pairs) {
  result <- permutation_test(data, group1 = pair[1], group2 = pair[2], nperm = nperm, stat = "median")
  comp_name <- paste(pair, collapse = " vs ")
  results[[comp_name]] <- result
  cat("Comparison", comp_name, "\n")
  cat("Observed difference (", pair[2], " - ", pair[1], "): ", result$obs_diff, "\n", sep = "")
  cat("P-value:", result$p_value, "\n\n")
}

# Set up plotting: 1 row, 3 columns with an outer margin for an overall title
par(mfrow = c(1, 3), oma = c(0, 0, 2, 0))
# Loop over the results to plot each histogram with improved x-axis limits
par(mfrow = c(1, 3), oma = c(0, 0, 2, 0))
render_threshold <- nperm * 0.0005  # Threshold below which a bin's count is considered too small

for(name in names(results)) {
  perm_diff <- results[[name]]$perm_diff
  obs_diff  <- results[[name]]$obs_diff
  
  # Compute histogram without plotting
  h <- hist(perm_diff, breaks = 50, plot = FALSE)
  
  # Identify bins where the count meets or exceeds the render threshold
  significant_bins <- which(h$counts >= render_threshold)
  
  if(length(significant_bins) > 0) {
    # Use the range of these bins as the basis for the x-axis
    new_xmin <- h$breaks[min(significant_bins)]
    new_xmax <- h$breaks[min(length(h$breaks), max(significant_bins) + 1)]
  } else {
    # If no bins meet the threshold, simply use the range of permuted values
    new_xmin <- min(perm_diff)
    new_xmax <- max(perm_diff)
  }
  
  # Ensure the observed statistic is within the limits
  new_xmin <- min(new_xmin, obs_diff)
  new_xmax <- max(new_xmax, obs_diff)
  
  # Add a 10% padding to the range
  padding <- (new_xmax - new_xmin) * 0.1
  xlim_val <- c(new_xmin - padding, new_xmax + padding)
  
  # Plot histogram with simplified title (e.g., "A vs B") and adjusted x-axis limits
  hist(perm_diff, breaks = 50,
       main = name,
       xlab = "Difference in Medians",
       col = "lightblue", border = "white", xlim = xlim_val)
  abline(v = obs_diff, col = "red", lwd = 2)
}

# Overall title for the combined plots
mtext("Controlled pairwise permutation distribution of differences between medians", outer = TRUE, cex = 1.5)

# Gather the p-values from each pairwise comparison into a vector
p_values <- sapply(results, function(x) x$p_value)
print("Original p-values:")
print(p_values)

# Adjust p-values using the Bonferroni correction
p_adjusted_bonf <- p.adjust(p_values, method = "bonferroni")
print("Bonferroni adjusted p-values:")
print(p_adjusted_bonf)

# Adjust p-values using the Benjamini-Hochberg procedure (FDR)
p_adjusted_fdr <- p.adjust(p_values, method = "BH")
print("FDR (Benjamini-Hochberg) adjusted p-values:")
print(p_adjusted_fdr)

# --- Initiation logic above this line ---
# --- Import logic below this line ---

# Load required libraries
library(readxl)   # Functions for importing data from Excel
library(janitor)  # Sanitizing column names
library(dplyr)
library(tidyr)    # For drop_na()
library(car)

# Prompt user to choose a file
file_path <- file.choose()

# Get available sheet names
sheet_names <- excel_sheets(file_path)

repeat {
  # Initialize an empty data frame for combining sheets
  combined_data <- data.frame()
  
  # Loop to allow combining data from multiple sheets
  repeat {
    # Print available sheet names each time a new sheet is requested
    cat("\nAvailable sheet names:\n")
    print(sheet_names)
    
    # Prompt for the sheet name until a valid one is entered
    repeat {
      sheet_name <- readline(prompt = "Enter the exact sheet name: ")
      if (sheet_name %in% sheet_names) break
      cat("Invalid sheet name. Try again.\n")
    }
    
    # Import and clean data from the chosen sheet
    temp_data <- read_excel(file_path, sheet = sheet_name) %>%
      clean_names() %>%  # Sanitize column names (e.g., "Treatment Group" → "treatment_group")
      mutate(across(where(is.character), ~ na_if(.x, "#N/A"))) %>%  # Replace "#N/A" with NA
      select(treatment_group, log_cfu_g_1)
    
    # Convert log_cfu_g_1 to numeric if applicable
    if (!is.numeric(temp_data$log_cfu_g_1)) {
      temp_data <- temp_data %>% mutate(log_cfu_g_1 = as.numeric(log_cfu_g_1))
    }
    
    # Combine data from this sheet with any previously imported sheets
    combined_data <- bind_rows(combined_data, temp_data)
    
    # Ask if the user wants to add another sheet from the same file
    add_sheet <- readline(prompt = "Would you like to add data from another sheet? (y/n): ")
    if (tolower(trimws(add_sheet)) != "y") break
  }
  
  # Use the combined data for further processing and analysis
  
  # Check for required columns
  required_columns <- c("treatment_group", "log_cfu_g_1")
  if (!all(required_columns %in% colnames(combined_data))) {
    stop("Required columns are missing. Please check the dataset.")
  }
  
  # Preview and summarize the data
  cat("Preview of the imported data:\n")
  print(head(combined_data))
  str(combined_data)
  summary(combined_data)
  
  # Drop rows with missing values in key columns
  combined_data <- combined_data %>% drop_na(treatment_group, log_cfu_g_1)
  
  # Ask user how to treat zero values before testing assumptions
  remove_zeroes <- readline(prompt = "Would you like to remove rows with zero values in log_cfu_g_1? (y/n): ")
  if (tolower(trimws(remove_zeroes)) == "y") {
    combined_data <- combined_data %>% filter(log_cfu_g_1 != 0)
    cat("Rows with zero values in log_cfu_g_1 have been removed.\n")
  } else {
    cat("Zero values in log_cfu_g_1 have been retained.\n")
  }
  
  # Convert treatment_group to factor
  combined_data <- combined_data %>% mutate(treatment_group = as.factor(treatment_group))
  
  # Summarize data by treatment group
  summary_stats <- combined_data %>%
    group_by(treatment_group) %>%
    summarise(
      mean_cfu = mean(log_cfu_g_1, na.rm = TRUE),
      sd_cfu = sd(log_cfu_g_1, na.rm = TRUE),
      median_cfu = median(log_cfu_g_1, na.rm = TRUE),
      n = n()
    )
  print(summary_stats)
  
  # ===== PERMUTATION TEST ON EXPERIMENTAL DATA =====
  # Rename columns to match those expected by the permutation_test() function
  exp_data <- combined_data %>% rename(Group = treatment_group, LogCFU = log_cfu_g_1)
  
  # Prompt user to select test statistic (median or mean)
  test_stat <- readline(prompt = "Enter test statistic to use (median/mean): ")
  test_stat <- tolower(trimws(test_stat))
  if (!test_stat %in% c("median", "mean")) {
    cat("Invalid test statistic selected. Defaulting to 'median'.\n")
    test_stat <- "median"
  }
  
  # Get unique groups from the experimental data
  groups <- unique(exp_data$Group)
  if (length(groups) < 2) {
    cat("Not enough groups for permutation test.\n")
  } else {
    # Create all pairwise comparisons
    group_pairs <- combn(as.character(groups), 2, simplify = FALSE)
    exp_results <- list()
    
    # Loop over each pair and perform the permutation test using the selected statistic
    for (pair in group_pairs) {
      result <- permutation_test(exp_data, group1 = pair[1], group2 = pair[2], nperm = nperm, stat = test_stat)
      comp_name <- paste(pair, collapse = " vs ")
      exp_results[[comp_name]] <- result
      cat("\nPermutation test for ", comp_name, "\n", sep = "")
      cat("Observed difference (", pair[2], " - ", pair[1], ") using ", test_stat, ": ", result$obs_diff, "\n", sep = "")
      cat("P-value: ", result$p_value, "\n", sep = "")
    }
    
    # Gather and adjust p-values for the experimental comparisons
    p_values_exp <- sapply(exp_results, function(x) x$p_value)
    cat("\nOriginal experimental p-values:\n")
    print(p_values_exp)
    cat("\nBonferroni adjusted experimental p-values:\n")
    print(p.adjust(p_values_exp, method = "bonferroni"))
    cat("\nFDR (Benjamini-Hochberg) adjusted experimental p-values:\n")
    print(p.adjust(p_values_exp, method = "BH"))
    cat("\nHolm adjusted experimental p-values:\n")
    print(p.adjust(p_values_exp, method = "holm"))
  }
  # ===== END OF PERMUTATION TEST ON EXPERIMENTAL DATA =====
  
  # Ask if the user wants to analyze the same file again
  reanalyze <- readline(prompt = "Would you like to analyze the same file again? (y/n): ")
  if (tolower(trimws(reanalyze)) != "y") {
    cat("Terminating the script.\n")
    break
  }
}

