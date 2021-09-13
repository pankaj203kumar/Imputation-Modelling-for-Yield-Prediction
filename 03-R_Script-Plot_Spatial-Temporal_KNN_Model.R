# THIS CODE RUNS KNN IMPUTATION MODEL WITH SPATIAL-TEMPORAL TRANSFERABILITY APPROACH
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

# READ THE GROUND INVENTORY & ASSOCIATED LIDAR METRIC FILES

metrics1_file <- read.xlsx("SET INPUT PLOT INTEGRATED 1st METRICS FILE IN .XLSX FORMAT") #PLOT_ID,X,Y,GROUND INVENTORY AND LIDAR DERIVED METRICS IN 1ST FILE
metrics2_file <- read.xlsx("SET INPUT PLOT INTEGRATED 2nd METRICS FILE IN .XLSX FORMAT") #PLOT_ID,X,Y,GROUND INVENTORY AND LIDAR DERIVED METRICS IN 2ND FILE

View(metrics1_file)
View(metrics2_file)

metrics1_lidar <- data.frame(cbind(metrics1_file[, 11], metrics1_file[, 78:160])) # COMBINE AGE METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_lidar)[1] <- "age" # PROVIDE COLUMN NAME TO THE AGE METRIC FOR 1ST
metrics2_lidar <- data.frame(cbind(metrics2_file[, 11], metrics2_file[, 78:160])) # COMBINE AGE METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_lidar)[1] <- "age" # PROVIDE COLUMN NAME TO THE AGE METRIC FOR 2ND

metrics1_GI <- data.frame(metrics1_file[, 2:10])# 1ST GROUND INVENTORY METRICS (STOCKING,BASAL AREA,HEIGHT,STEM VOLUME,...
# ...TOTAL RECOVERABLE VOLUME,SAW,INDUSTRIAL,CHIP,PULP VARIABLES)
metrics2_GI <- data.frame(metrics2_file[, 2:10])# 2ND GROUND INVENTORY METRICS (STOCKING,BASAL AREA,HEIGHT,STEM VOLUME,...
# ...TOTAL RECOVERABLE VOLUME,SAW,INDUSTRIAL,CHIP,PULP VARIABLES)

metrics1_GI_lidar <- data.frame(cbind(metrics1_GI, metrics1_lidar)) # COMBINE 1ST GROUND INVENTORY AND LIDAR METRICS
metrics2_GI_lidar <- data.frame(cbind(metrics2_GI, metrics2_lidar)) # COMBINE 2ND GROUND INVENTORY AND LIDAR METRICS

# DEVELOP THE SPATIAL-TEMPORAL KNN MODEL AND IMPUTE THE GROUND INVENTORY METRICS
set.seed(1)
metrics1_lidar_metrics2 <- data.frame(rbind(metrics1_lidar, metrics2_lidar)) # COMBINE 1ST AND 2ND LIDAR METRICS
metrics1_knn_metrics2 <- yai(x = metrics1_lidar_metrics2, y = metrics1_GI, k = 1, noRefs = TRUE, method = "randomForest",
                             mtry = floor(ncol(metrics1_lidar_metrics2)/3), ntree = 2000, rfMode = "regression") # KNN MODEL IS...
# ...DEVELOPED FROM 1ST METRICS TO IMPUTE 2ND GROUND INVENTORY METRICS
yaiVarImp(metrics1_knn_metrics2, nTop = 10, col = "gray", plot = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_GI_imp_metrics1 <- impute(metrics1_knn_metrics2, ancillaryData = data.frame(rbind(metrics1_GI_lidar, metrics2_GI_lidar)),
                                   method = "mean", vars = yvars(metrics1_knn_metrics2)) #2ND GROUND INVENTORY METRICS ARE...
# ...IMPUTED FROM 1ST KNN MODEL
plot(metrics2_GI_imp_metrics1) # PLOT IMPUTED 2ND GROUND INVENTORY METRICS

metrics2_lidar_metrics1 <- data.frame(rbind(metrics2_lidar, metrics1_lidar)) # COMBINE 2ND AND 1ST LIDAR METRICS
metrics2_knn_metrics1 <- yai(x = metrics2_lidar_metrics1, y = metrics2_GI, k = 1, noRefs = TRUE, method = "randomForest",
                             mtry = floor(ncol(metrics2_lidar_metrics1)/3), ntree = 2000, rfMode = "regression") # KNN MODEL IS...
# ...DEVELOPED FROM 2ND METRICS TO IMPUTE GROUND INVENTORY 1ST METRICS
yaiVarImp(metrics2_knn_metrics1, nTop = 10, col = "gray", plot = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_GI_imp_metrics2 <- impute(metrics2_knn_metrics1, ancillaryData = data.frame(rbind(metrics2_GI_lidar, metrics1_GI_lidar)),
                                   method = "mean", vars = yvars(metrics2_knn_metrics1)) #1ST GROUND INVENTORY METRICS ARE...
# ...IMPUTED FROM 2ND KNN MODEL
plot(metrics1_GI_imp_metrics2) # PLOT IMPUTED 1ST GROUND INVENTORY METRICS

# ACCURACY ASSESSMENT OF IMPUTED 1ST AND 2ND GROUND INVENTORY METRICS WITH RESPECT TO THEIR ORIGINAL METRIC VALUES...
# ...BASED ON THE ESTIMATION OF NORMALIZED RMSE (NRMSE) AND PERCENT BIASNESS (PBIAS) MEASURES

set.seed(2)
metrics1_n_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 1], metrics1_GI_imp_metrics2[, 10], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_ba_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 2], metrics1_GI_imp_metrics2[, 11], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_h_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 3], metrics1_GI_imp_metrics2[, 12], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_vs_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 4], metrics1_GI_imp_metrics2[, 13], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_trv_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 5], metrics1_GI_imp_metrics2[, 14], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_saw_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 6], metrics1_GI_imp_metrics2[, 15], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_industrial_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 7], metrics1_GI_imp_metrics2[, 16], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_chip_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 8], metrics1_GI_imp_metrics2[, 17], na.rm = TRUE, norm = "maxmin"), 3)
metrics1_pulp_nrmse_metrics2 <- round(nrmse(metrics1_GI_imp_metrics2[, 9], metrics1_GI_imp_metrics2[, 18], na.rm = TRUE, norm = "maxmin"), 3)

metrics1_n_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 1], metrics1_GI_imp_metrics2[, 10], na.rm = TRUE), 3)
metrics1_ba_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 2], metrics1_GI_imp_metrics2[, 11], na.rm = TRUE), 3)
metrics1_h_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 3], metrics1_GI_imp_metrics2[, 12], na.rm = TRUE), 3)
metrics1_vs_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 4], metrics1_GI_imp_metrics2[, 13], na.rm = TRUE), 3)
metrics1_trv_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 5], metrics1_GI_imp_metrics2[, 14], na.rm = TRUE), 3)
metrics1_saw_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 6], metrics1_GI_imp_metrics2[, 15], na.rm = TRUE), 3)
metrics1_industrial_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 7], metrics1_GI_imp_metrics2[, 16], na.rm = TRUE), 3)
metrics1_chip_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 8], metrics1_GI_imp_metrics2[, 17], na.rm = TRUE), 3)
metrics1_pulp_pbias_metrics2 <- round(pbias(metrics1_GI_imp_metrics2[, 9], metrics1_GI_imp_metrics2[, 18], na.rm = TRUE), 3)

metrics2_n_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 1], metrics2_GI_imp_metrics1[, 10], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_ba_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 2], metrics2_GI_imp_metrics1[, 11], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_h_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 3], metrics2_GI_imp_metrics1[, 12], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_vs_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 4], metrics2_GI_imp_metrics1[, 13], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_trv_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 5], metrics2_GI_imp_metrics1[, 14], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_saw_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 6], metrics2_GI_imp_metrics1[, 15], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_industrial_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 7], metrics2_GI_imp_metrics1[, 16], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_chip_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 8], metrics2_GI_imp_metrics1[, 17], na.rm = TRUE, norm = "maxmin"), 3)
metrics2_pulp_nrmse_metrics1 <- round(nrmse(metrics2_GI_imp_metrics1[, 9], metrics2_GI_imp_metrics1[, 18], na.rm = TRUE, norm = "maxmin"), 3)

metrics2_n_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 1], metrics2_GI_imp_metrics1[, 10], na.rm = TRUE), 3)
metrics2_ba_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 2], metrics2_GI_imp_metrics1[, 11], na.rm = TRUE), 3)
metrics2_h_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 3], metrics2_GI_imp_metrics1[, 12], na.rm = TRUE), 3)
metrics2_vs_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 4], metrics2_GI_imp_metrics1[, 13], na.rm = TRUE), 3)
metrics2_trv_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 5], metrics2_GI_imp_metrics1[, 14], na.rm = TRUE), 3)
metrics2_saw_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 6], metrics2_GI_imp_metrics1[, 15], na.rm = TRUE), 3)
metrics2_industrial_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 7], metrics2_GI_imp_metrics1[, 16], na.rm = TRUE), 3)
metrics2_chip_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 8], metrics2_GI_imp_metrics1[, 17], na.rm = TRUE), 3)
metrics2_pulp_pbias_metrics1 <- round(pbias(metrics2_GI_imp_metrics1[, 9], metrics2_GI_imp_metrics1[, 18], na.rm = TRUE), 3)

# COMBINE ESTIMATED NRMSE AND PBIAS ACCURACY MEASURES FOR THE 1ST AND 2ND GROUND INVENTORY METRICS
set.seed(3)
metrics1_GI_accuracy_metrics2 <- cbind(rbind("Variables", "Stocking", "Basal Area", "Height", "Stem Volume",
                                             "Total Recover Volume", "Saw", "Industrial", "Chip", "Pulp"), 
                                       rbind("sRMSE",
                                             metrics1_n_nrmse_metrics2,
                                             metrics1_ba_nrmse_metrics2,
                                             metrics1_h_nrmse_metrics2,
                                             metrics1_vs_nrmse_metrics2,
                                             metrics1_trv_nrmse_metrics2,
                                             metrics1_saw_nrmse_metrics2,
                                             metrics1_industrial_nrmse_metrics2,
                                             metrics1_chip_nrmse_metrics2,
                                             metrics1_pulp_nrmse_metrics2),
                                       rbind("pBias",
                                             metrics1_n_pbias_metrics2,
                                             metrics1_ba_pbias_metrics2,
                                             metrics1_h_pbias_metrics2,
                                             metrics1_vs_pbias_metrics2,
                                             metrics1_trv_pbias_metrics2,
                                             metrics1_saw_pbias_metrics2,
                                             metrics1_industrial_pbias_metrics2,
                                             metrics1_chip_pbias_metrics2,
                                             metrics1_pulp_pbias_metrics2))

metrics2_GI_accuracy_metrics1 <- cbind(rbind("Variables", "Stocking", "Basal Area", "Height", "Stem Volume",
                                             "Total Recover Volume", "Saw", "Industrial", "Chip", "Pulp"), 
                              rbind("sRMSE",
                                    metrics2_n_nrmse_metrics1,
                                    metrics2_ba_nrmse_metrics1,
                                    metrics2_h_nrmse_metrics1,
                                    metrics2_vs_nrmse_metrics1,
                                    metrics2_trv_nrmse_metrics1,
                                    metrics2_saw_nrmse_metrics1,
                                    metrics2_industrial_nrmse_metrics1,
                                    metrics2_chip_nrmse_metrics1,
                                    metrics2_pulp_nrmse_metrics1),
                              rbind("pBias",
                                    metrics2_n_pbias_metrics1,
                                    metrics2_ba_pbias_metrics1,
                                    metrics2_h_pbias_metrics1,
                                    metrics2_vs_pbias_metrics1,
                                    metrics2_trv_pbias_metrics1,
                                    metrics2_saw_pbias_metrics1,
                                    metrics2_industrial_pbias_metrics1,
                                    metrics2_chip_pbias_metrics1,
                                    metrics2_pulp_pbias_metrics1))





# WRITE ORIGINAL AND IMPUTED 1ST AND 2ND GROUND INVENTORY METRICS ALONG WITH NRMSE AND PBIAS ACCURACY MEASURES IN .XLSX FORMAT
set.seed(4)
metrics1_GI_wb_metrics2 <- createWorkbook("metrics1_GI_metrics2")
addWorksheet(metrics1_GI_wb_metrics2, "Imputation_Results")
writeData(metrics1_GI_wb_metrics2, sheet = "Imputation_Results", metrics1_GI_imp_metrics2)
addWorksheet(metrics1_GI_wb_metrics2, "Accuracy_Results")
writeData(metrics1_GI_wb_metrics2, sheet = "Accuracy_Results", metrics1_GI_accuracy_metrics2)
addWorksheet(metrics1_GI_wb_metrics2, "Imputation_Figures")
saveWorkbook(metrics1_GI_wb_metrics2,
             "WRITE FOLDER PATH AND OUTPUT PLOT BASED 1ST FILE IN .XLSX FORMAT",
             overwrite = TRUE)

metrics2_GI_wb_metrics1 <- createWorkbook("metrics2_GI_metrics1")
addWorksheet(metrics2_GI_wb_metrics1, "Imputation_Results")
writeData(metrics2_GI_wb_metrics1, sheet = "Imputation_Results", metrics2_GI_imp_metrics1)
addWorksheet(metrics2_GI_wb_metrics1, "Accuracy_Results")
writeData(metrics2_GI_wb_metrics1, sheet = "Accuracy_Results", metrics2_GI_accuracy_metrics1)
addWorksheet(metrics2_GI_wb_metrics1, "Imputation_Figures")
saveWorkbook(metrics2_GI_wb_metrics1,
             "WRITE FOLDER PATH AND OUTPUT PLOT BASED 2ND FILE IN .XLSX FORMAT",
             overwrite = TRUE)

