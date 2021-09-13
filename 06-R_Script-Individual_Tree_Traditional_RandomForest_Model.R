# THIS CODE RUNS INDIVIDUAL TREE BASED RANDOM FOREST IMPUTATION MODEL WITH TRADITIONAL (LOOCV-LEAVE ONE OUT CROSS VALIDATION) APPROACH
# -----------------------------------------------

rm(list = ls()) #REMOVES ALL THE VARIABLES
cat("\f") #CLEAR SCREEN

setwd("SET WORKING DIRECTORY")

require(openxlsx)
require(matlib)
require(Metrics)
require(ggplot2)
require(randomForest)
require(hydroGOF)

# READ THE GROUND INVENTORY & ASSOCIATED LIDAR METRICS FILES

metrics_file <- read.xlsx("SET INPUT INDIVIDUAL TREE BASED INTEGRATED METRICS FILE IN .XLSX FORMAT") #PLOT_ID,X,Y,GROUND INVENTORY AND LIDAR DERIVED METRICS

View(metrics_file)

# DEVELOP THE TRADITIONAL RANDOM FOREST MODELS & ACCURACY ASSESSMENT FOR INDIVIDUAL DATASET
set.seed(1)
start.time = Sys.time()

metrics_lidar <- data.frame(cbind(metrics_file[, 11], metrics_file[, 12:228])) # COMBINE AGE METRIC WITH THE LIDAR METRICS
colnames(metrics_lidar)[1] <- "age" # PROVIDE COLUMN NAME TO THE AGE METRIC
metrics_ba_lidar <- data.frame(cbind(metrics_file[, 3], metrics_lidar)) # COMBINE BASAL AREA METRIC WITH THE LIDAR METRICS
colnames(metrics_ba_lidar)[1] <- "ba" # PROVIDE COLUMN NAME TO THE BASAL AREA METRIC
metrics_h_lidar <- data.frame(cbind(metrics_file[, 4], metrics_lidar)) # COMBINE HEIGHT METRIC WITH THE LIDAR METRICS
colnames(metrics_h_lidar)[1] <- "h" # PROVIDE COLUMN NAME TO THE HEIGHT METRIC
metrics_vs_lidar <- data.frame(cbind(metrics_file[, 5], metrics_lidar)) # COMBINE STEM VOLUME METRIC WITH THE LIDAR METRICS
colnames(metrics_vs_lidar)[1] <- "vs" # PROVIDE COLUMN NAME TO THE STEM VOLUME METRIC
metrics_trv_lidar <- data.frame(cbind(metrics_file[, 6], metrics_lidar)) # COMBINE TOTAL RECOVERABLE VOLUME METRIC WITH THE LIDAR METRICS
colnames(metrics_trv_lidar)[1] <- "trv" # PROVIDE COLUMN NAME TO THE TOTAL RECOVERABLE VOLUME METRIC
metrics_saw_lidar <- data.frame(cbind(metrics_file[, 7], metrics_lidar)) # COMBINE SAW METRIC WITH THE LIDAR METRICS
colnames(metrics_saw_lidar)[1] <- "saw" # PROVIDE COLUMN NAME TO THE SAW METRIC
metrics_industrial_lidar <- data.frame(cbind(metrics_file[, 8], metrics_lidar)) # COMBINE INDUSTRIAL METRIC WITH THE LIDAR METRICS
colnames(metrics_industrial_lidar)[1] <- "industrial" # PROVIDE COLUMN NAME TO THE INDUSTRIAL METRIC
metrics_chip_lidar <- data.frame(cbind(metrics_file[, 9], metrics_lidar)) # COMBINE CHIP METRIC WITH THE LIDAR METRICS
colnames(metrics_chip_lidar)[1] <- "chip" # PROVIDE COLUMN NAME TO THE CHIP METRIC
metrics_pulp_lidar <- data.frame(cbind(metrics_file[, 10], metrics_lidar)) # COMBINE PULP METRIC WITH THE LIDAR METRICS
colnames(metrics_pulp_lidar)[1] <- "pulp" # PROVIDE COLUMN NAME TO THE PULP METRIC

metrics_ba <- metrics_ba_lidar$ba # BASAL AREA METRIC
metrics_h <- metrics_h_lidar$h # HEIGHT METRIC
metrics_vs <- metrics_vs_lidar$vs # STEM VOLUME METRIC
metrics_trv <- metrics_trv_lidar$trv # TOTAL RECOVERABLE VOLUME METRIC
metrics_saw <- metrics_saw_lidar$saw # SAW METRIC
metrics_industrial <- metrics_industrial_lidar$industrial # INDUSTRIAL METRIC
metrics_chip <- metrics_chip_lidar$chip # CHIP METRIC
metrics_pulp <- metrics_pulp_lidar$pulp # PULP METRIC

nrow_ind <- nrow(metrics_file)
metrics_ba_imp <- data.frame()
metrics_h_imp <- data.frame()
metrics_vs_imp <- data.frame()
metrics_trv_imp <- data.frame()
metrics_saw_imp <- data.frame()
metrics_industrial_imp <- data.frame()
metrics_chip_imp <- data.frame()
metrics_pulp_imp <- data.frame()

pb <- winProgressBar(title = "Progress Bar", min = 1, max = nrow_ind, initial = 1, width = 300)
for (m in 1:nrow_ind)
{
  metrics_ba_lidar_in <- metrics_ba_lidar[-c(m),] # ONE BASAL AREA BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_ba_rf <- randomForest(ba~., data = metrics_ba_lidar_in, ntree = 100, mtry = floor(ncol(metrics_ba_lidar_in)/3),
                                importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_ba_imp_out <- predict(metrics_ba_rf, metrics_lidar[m,]) # BASAL AREA METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_ba_imp <- rbind(metrics_ba_imp, metrics_ba_imp_out) #BASAL AREA METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
  metrics_h_lidar_in <- metrics_h_lidar[-c(m),] # ONE HEIGHT BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_h_rf <- randomForest(h~., data = metrics_h_lidar_in, ntree = 100, mtry = floor(ncol(metrics_h_lidar_in)/3),
                               importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_h_imp_out <- predict(metrics_h_rf, metrics_lidar[m,]) # HEIGHT METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_h_imp <- rbind(metrics_h_imp, metrics_h_imp_out) #HEIGHT METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
  metrics_vs_lidar_in <- metrics_vs_lidar[-c(m),] # ONE STEM VOLUME BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_vs_rf <- randomForest(vs~., data = metrics_vs_lidar_in, ntree = 100, mtry = floor(ncol(metrics_vs_lidar_in)/3),
                                importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_vs_imp_out <- predict(metrics_vs_rf, metrics_lidar[m,]) # STEM VOLUME METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_vs_imp <- rbind(metrics_vs_imp, metrics_vs_imp_out) #STEM VOLUME METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
  metrics_trv_lidar_in <- metrics_trv_lidar[-c(m),] # ONE TOTAL RECOVERABLE VOLUME BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_trv_rf <- randomForest(trv~., data = metrics_trv_lidar_in, ntree = 100, mtry = floor(ncol(metrics_trv_lidar_in)/3),
                                 importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_trv_imp_out <- predict(metrics_trv_rf, metrics_lidar[m,]) # TOTAL RECOVERABLE VOLUME METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_trv_imp <- rbind(metrics_trv_imp, metrics_trv_imp_out) #TOTAL RECOVERABLE VOLUME METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
  metrics_saw_lidar_in <- metrics_saw_lidar[-c(m),] # ONE SAW BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_saw_rf <- randomForest(saw~., data = metrics_saw_lidar_in, ntree = 100, mtry = floor(ncol(metrics_saw_lidar_in)/3),
                                 importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_saw_imp_out <- predict(metrics_saw_rf, metrics_lidar[m,]) # SAW METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_saw_imp <- rbind(metrics_saw_imp, metrics_saw_imp_out) #SAW METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
  metrics_industrial_lidar_in <- metrics_industrial_lidar[-c(m),] # ONE INDUSTRIAL BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_industrial_rf <- randomForest(industrial~., data = metrics_industrial_lidar_in, ntree = 100, mtry = floor(ncol(metrics_industrial_lidar_in)/3),
                                        importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_industrial_imp_out <- predict(metrics_industrial_rf, metrics_lidar[m,]) # INDUSTRIAL METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_industrial_imp <- rbind(metrics_industrial_imp, metrics_industrial_imp_out) #INDUSTRIAL METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
  metrics_chip_lidar_in <- metrics_chip_lidar[-c(m),] # ONE CHIP BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_chip_rf <- randomForest(chip~., data = metrics_chip_lidar_in, ntree = 100, mtry = floor(ncol(metrics_chip_lidar_in)/3),
                                  importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_chip_imp_out <- predict(metrics_chip_rf, metrics_lidar[m,]) # CHIP METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_chip_imp <- rbind(metrics_chip_imp, metrics_chip_imp_out) #CHIP METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
  metrics_pulp_lidar_in <- metrics_pulp_lidar[-c(m),] # ONE PULP BASED PLOT SAMPLE IS LEFT-OUT DURING EACH ITERATION
  metrics_pulp_rf <- randomForest(pulp~., data = metrics_pulp_lidar_in, ntree = 100, mtry = floor(ncol(metrics_pulp_lidar_in)/3),
                                  importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM REST OF THE PLOT SAMPLES
  metrics_pulp_imp_out <- predict(metrics_pulp_rf, metrics_lidar[m,]) # PULP METRIC FOR THE LEFT-OUT SAMPLE IS IMPUTED FROM THE MODEL
  metrics_pulp_imp <- rbind(metrics_pulp_imp, metrics_pulp_imp_out) #PULP METRIC FOR EACH LEFT-OUT SAMPLE IS COMBINED
  
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
metrics_ba_nrmse <- round(nrmse(metrics_ba_imp, metrics_ba, na.rm = TRUE, norm = "maxmin"), 3)
metrics_h_nrmse <- round(nrmse(metrics_h_imp, metrics_h, na.rm = TRUE, norm = "maxmin"), 3)
metrics_vs_nrmse <- round(nrmse(metrics_vs_imp, metrics_vs, na.rm = TRUE, norm = "maxmin"), 3)
metrics_trv_nrmse <- round(nrmse(metrics_trv_imp, metrics_trv, na.rm = TRUE, norm = "maxmin"), 3)
metrics_saw_nrmse <- round(nrmse(metrics_saw_imp, metrics_saw, na.rm = TRUE, norm = "maxmin"), 3)
metrics_industrial_nrmse <- round(nrmse(metrics_industrial_imp, metrics_industrial, na.rm = TRUE, norm = "maxmin"), 3)
metrics_chip_nrmse <- round(nrmse(metrics_chip_imp, metrics_chip, na.rm = TRUE, norm = "maxmin"), 3)
metrics_pulp_nrmse <- round(nrmse(metrics_pulp_imp, metrics_pulp, na.rm = TRUE, norm = "maxmin"), 3)

metrics_ba_pbias <- round(pbias(metrics_ba_imp, metrics_ba, na.rm = TRUE), 3)
metrics_h_pbias <- round(pbias(metrics_h_imp, metrics_h, na.rm = TRUE), 3)
metrics_vs_pbias <- round(pbias(metrics_vs_imp, metrics_vs, na.rm = TRUE), 3)
metrics_trv_pbias <- round(pbias(metrics_trv_imp, metrics_trv, na.rm = TRUE), 3)
metrics_saw_pbias <- round(pbias(metrics_saw_imp, metrics_saw, na.rm = TRUE), 3)
metrics_industrial_pbias <- round(pbias(metrics_industrial_imp, metrics_industrial, na.rm = TRUE), 3)
metrics_chip_pbias <- round(pbias(metrics_chip_imp, metrics_chip, na.rm = TRUE), 3)
metrics_pulp_pbias <- round(pbias(metrics_pulp_imp, metrics_pulp, na.rm = TRUE), 3)

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

metrics_GI_imp <- cbind(metrics_ba_imp, metrics_h_imp, metrics_vs_imp, metrics_trv_imp,
                    metrics_saw_imp, metrics_industrial_imp, metrics_chip_imp, metrics_pulp_imp,
                    metrics_ba, metrics_h, metrics_vs, metrics_trv,
                    metrics_saw, metrics_industrial, metrics_chip, metrics_pulp)

colnames(metrics_GI_imp)[1] <- "ba"
colnames(metrics_GI_imp)[2] <- "h"
colnames(metrics_GI_imp)[3] <- "vs"
colnames(metrics_GI_imp)[4] <- "trv"
colnames(metrics_GI_imp)[5] <- "saw"
colnames(metrics_GI_imp)[6] <- "industrial"
colnames(metrics_GI_imp)[7] <- "chip"
colnames(metrics_GI_imp)[8] <- "pulp"

# WRITE ORIGINAL AND IMPUTED GROUND INVENTORY METRICS ALONG WITH NRMSE AND PBIAS ACCURACY MEASURES IN .XLSX FORMAT
metrics_GI_wb <- createWorkbook("metrics_GI")
addWorksheet(metrics_GI_wb, "Imputation_Results")
writeData(metrics_GI_wb, sheet = "Imputation_Results", metrics_GI_imp)
addWorksheet(metrics_GI_wb, "Accuracy_Results")
writeData(metrics_GI_wb, sheet = "Accuracy_Results", metrics_GI_accuracy)
saveWorkbook(metrics_GI_wb,
             "WRITE FOLDER PATH AND OUTPUT INDIVIDUAL TREE BASED FILE IN .XLSX FORMAT",
             overwrite = TRUE)
