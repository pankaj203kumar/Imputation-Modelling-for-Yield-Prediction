# THIS CODE RUNS INDIVIDUAL TREE BASED KNN IMPUTATION MODEL WITH TRADITIONAL (LOOCV-LEAVE ONE OUT CROSS VALIDATION) APPROACH
# -----------------------------------------------

rm(list = ls()) #REMOVES ALL THE VARIABLES
cat("\f") #CLEAR SCREEN

setwd("SET WORKING DIRECTORY")

require(openxlsx)
require(matlib)
require(Metrics)
require(ggplot2)
require(yaImpute)
require(randomForest)
require(hydroGOF)

# READ THE GROUND INVENTORY & ASSOCIATED LIDAR METRICS FILE

metrics_file <- read.xlsx("SET INPUT INDIVIDUAL TREE BASED INTEGRATED METRICS FILE IN .XLSX FORMAT") #PLOT_ID,Tree_ID,GROUND INVENTORY AND LIDAR DERIVED METRICS

View(metrics_file)

# DEVELOP THE TRADITIONAL KNN MODEL AND IMPUTE THE GROUND INVENTORY METRICS
set.seed(1)
start.time = Sys.time()

metrics_lidar <- data.frame(cbind(metrics_file[, 11], metrics_file[, 12:228])) # COMBINE AGE METRIC WITH THE LIDAR METRICS
colnames(metrics_lidar)[1] <- "age" # PROVIDE COLUMN NAME TO THE AGE METRIC
metrics_GI <- data.frame(metrics_file[, 3:10]) # GROUND INVENTORY METRICS (BASAL AREA,HEIGHT,STEM VOLUME,...
# ...TOTAL RECOVERABLE VOLUME,SAW,INDUSTRIAL,CHIP,PULP VARIABLES)
metrics_GI_imp <- data.frame()
nrow_ind <- nrow(metrics_file)

pb <- winProgressBar(title = "Progress Bar", min = 1, max = nrow_ind, initial= 1, width = 300)
for (m in 1:nrow_ind)
{
  metrics_GI_in <- metrics_GI[-c(m),] # ONE PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_knn <- yai(x = metrics_lidar, y = metrics_GI_in, k = 1, noRefs = TRUE, method = "randomForest",
                     mtry = floor(ncol(metrics_lidar)/3), ntree = 2000, rfMode = "regression") # KNN MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_GI_imp_out <- impute(metrics_knn, ancillaryData = data.frame(metrics_file[, 3:228]),
                               method = "mean", vars = yvars(metrics_knn)) # GROUND INVENTORY METRICS FOR THE LEFT-OUT SAMPLE ARE IMPUTED FROM THE MODEL
  
  metrics_GI_imp <- rbind(metrics_GI_imp, metrics_GI_imp_out) #GROUND INVENTORY METRICS FOR EACH LEFT-OUT SAMPLE ARE COMBINED
  Sys.sleep(0.1)
  setWinProgressBar(pb, m, title = paste(round((m * 100)/nrow_ind, 0), "% done"))
}
close(pb)
end.time = Sys.time()
time.taken = end.time - start.time
print(time.taken)

# ACCURACY ASSESSMENT OF IMPUTED GROUND INVENTORY METRICS WITH RESPECT TO THEIR ORIGINAL METRIC VALUES...
# ...BASED ON THE ESTIMATION OF NORMALIZED RMSE (NRMSE) AND PERCENT BIASNESS (PBIAS) MEASURES
set.seed(2)
metrics_ba_nrmse <- round(nrmse(metrics_GI_imp[, 1], metrics_GI_imp[, 9], na.rm = TRUE, norm = "maxmin"), 3)
metrics_h_nrmse <- round(nrmse(metrics_GI_imp[, 2], metrics_GI_imp[, 10], na.rm = TRUE, norm = "maxmin"), 3)
metrics_vs_nrmse <- round(nrmse(metrics_GI_imp[, 3], metrics_GI_imp[, 11], na.rm = TRUE, norm = "maxmin"), 3)
metrics_trv_nrmse <- round(nrmse(metrics_GI_imp[, 4], metrics_GI_imp[, 12], na.rm = TRUE, norm = "maxmin"), 3)
metrics_saw_nrmse <- round(nrmse(metrics_GI_imp[, 5], metrics_GI_imp[, 13], na.rm = TRUE, norm = "maxmin"), 3)
metrics_industrial_nrmse <- round(nrmse(metrics_GI_imp[, 6], metrics_GI_imp[, 14], na.rm = TRUE, norm = "maxmin"), 3)
metrics_chip_nrmse <- round(nrmse(metrics_GI_imp[, 7], metrics_GI_imp[, 15], na.rm = TRUE, norm = "maxmin"), 3)
metrics_pulp_nrmse <- round(nrmse(metrics_GI_imp[, 8], metrics_GI_imp[, 16], na.rm = TRUE, norm = "maxmin"), 3)

metrics_ba_pbias <- round(pbias(metrics_GI_imp[, 1], metrics_GI_imp[, 9], na.rm = TRUE), 3)
metrics_h_pbias <- round(pbias(metrics_GI_imp[, 2], metrics_GI_imp[, 10], na.rm = TRUE), 3)
metrics_vs_pbias <- round(pbias(metrics_GI_imp[, 3], metrics_GI_imp[, 11], na.rm = TRUE), 3)
metrics_trv_pbias <- round(pbias(metrics_GI_imp[, 4], metrics_GI_imp[, 12], na.rm = TRUE), 3)
metrics_saw_pbias <- round(pbias(metrics_GI_imp[, 5], metrics_GI_imp[, 13], na.rm = TRUE), 3)
metrics_industrial_pbias <- round(pbias(metrics_GI_imp[, 6], metrics_GI_imp[, 14], na.rm = TRUE), 3)
metrics_chip_pbias <- round(pbias(metrics_GI_imp[, 7], metrics_GI_imp[, 15], na.rm = TRUE), 3)
metrics_pulp_pbias <- round(pbias(metrics_GI_imp[, 8], metrics_GI_imp[, 16], na.rm = TRUE), 3)

# COMBINE ESTIMATED NRMSE AND PBIAS ACCURACY MEASURES FOR THE GROUND INVENTORY METRICS
set.seed(3)
metrics_GI_accuracy <- cbind(rbind("Variables", "Basal Area", "Height", "Stem Volume", "Total R. Volume",
                               "Saw", "Industrial", "Chip", "Pulp"), 
                                 rbind("sRMSE",
                                       metrics_ba_nrmse,
                                       metrics_h_nrmse,
                                       metrics_vs_nrmse,
                                       metrics_trv_nrmse,
                                       metrics_saw_nrmse,
                                       metrics_industrial_nrmse,
                                       metrics_chip_nrmse,
                                       metrics_pulp_nrmse),
                                 rbind("pBias",
                                       metrics_ba_pbias,
                                       metrics_h_pbias,
                                       metrics_vs_pbias,
                                       metrics_trv_pbias,
                                       metrics_saw_pbias,
                                       metrics_industrial_pbias,
                                       metrics_chip_pbias,
                                       metrics_pulp_pbias))

# WRITE ORIGINAL AND IMPUTED GROUND INVENTORY METRICS ALONG WITH NRMSE AND PBIAS ACCURACY MEASURES IN .XLSX FORMAT
set.seed(4)
metrics_GI_wb <- createWorkbook("metrics_GI")
addWorksheet(metrics_GI_wb, "Imputation_Results")
writeData(metrics_GI_wb, sheet = "Imputation_Results", metrics_GI_imp)
addWorksheet(metrics_GI_wb, "Accuracy_Results")
writeData(metrics_GI_wb, sheet = "Accuracy_Results", metrics_GI_accuracy)
saveWorkbook(metrics_GI_wb,
             "WRITE FOLDER PATH AND OUTPUT INDIVIDUAL TREE BASED FILE IN .XLSX FORMAT",
             overwrite = TRUE)
