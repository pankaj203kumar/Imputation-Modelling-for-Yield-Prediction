# THIS CODE RUNS RANDOM FOREST IMPUTATION MODEL WITH SPATIAL-TEMPORAL TRANSFERABILITY APPROACH
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

# READ THE GROUND INVENTORY & ASSOCIATED LIDAR METRIC FILES

metrics1_file <- read.xlsx("SET INPUT PLOT INTEGRATED 1st METRICS FILE IN .XLSX FORMAT") #PLOT_ID,X,Y,GROUND INVENTORY AND LIDAR DERIVED METRICS IN 1ST FILE
metrics2_file <- read.xlsx("SET INPUT PLOT INTEGRATED 2nd METRICS FILE IN .XLSX FORMAT") #PLOT_ID,X,Y,GROUND INVENTORY AND LIDAR DERIVED METRICS IN 2ND FILE

View(metrics1_file)
View(metrics2_file)

metrics1_lidar <- data.frame(cbind(metrics1_file[, 11], metrics1_file[, 12:160])) # COMBINE AGE METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_lidar)[1] <- "age" # PROVIDE COLUMN NAME TO THE AGE METRIC FOR 1ST
metrics1_n_lidar <- data.frame(cbind(metrics1_file[, 2], metrics1_lidar)) # COMBINE STOCKING METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_n_lidar)[1] <- "n" # PROVIDE COLUMN NAME TO THE STOCKING METRIC FOR 1ST
metrics1_ba_lidar <- data.frame(cbind(metrics1_file[, 3], metrics1_lidar)) # COMBINE BASAL AREA METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_ba_lidar)[1] <- "ba" # PROVIDE COLUMN NAME TO THE BASAL AREA METRIC FOR 1ST
metrics1_h_lidar <- data.frame(cbind(metrics1_file[, 4], metrics1_lidar)) # COMBINE HEIGHT METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_h_lidar)[1] <- "h" # PROVIDE COLUMN NAME TO THE HEIGHT METRIC FOR 1ST
metrics1_vs_lidar <- data.frame(cbind(metrics1_file[, 5], metrics1_lidar)) # COMBINE STEM VOLUME METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_vs_lidar)[1] <- "vs" # PROVIDE COLUMN NAME TO THE STEM VOLUME METRIC FOR 1ST
metrics1_trv_lidar <- data.frame(cbind(metrics1_file[, 6], metrics1_lidar)) # COMBINE TOTAL RECOVERABLE VOLUME METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_trv_lidar)[1] <- "trv" # PROVIDE COLUMN NAME TO THE TOTAL RECOVERABLE VOLUME METRIC FOR 1ST
metrics1_saw_lidar <- data.frame(cbind(metrics1_file[, 7], metrics1_lidar)) # COMBINE SAW METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_saw_lidar)[1] <- "saw" # PROVIDE COLUMN NAME TO THE SAW METRIC FOR 1ST
metrics1_industrial_lidar <- data.frame(cbind(metrics1_file[, 8], metrics1_lidar)) # COMBINE INDUSTRIAL METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_industrial_lidar)[1] <- "industrial" # PROVIDE COLUMN NAME TO THE INDUSTRIAL METRIC FOR 1ST
metrics1_chip_lidar <- data.frame(cbind(metrics1_file[, 9], metrics1_lidar)) # COMBINE CHIP METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_chip_lidar)[1] <- "chip" # PROVIDE COLUMN NAME TO THE CHIP METRIC FOR 1ST
metrics1_pulp_lidar <- data.frame(cbind(metrics1_file[, 10], metrics1_lidar)) # COMBINE PULP METRIC WITH THE LIDAR METRICS FOR 1ST
colnames(metrics1_pulp_lidar)[1] <- "pulp" # PROVIDE COLUMN NAME TO THE PULP METRIC FOR 1ST

metrics1_n <- metrics1_n_lidar$n # 1ST STOCKING METRIC
metrics1_ba <- metrics1_ba_lidar$ba # 1ST BASAL AREA METRIC
metrics1_h <- metrics1_h_lidar$h # 1ST HEIGHT METRIC
metrics1_vs <- metrics1_vs_lidar$vs # 1ST STEM VOLUME METRIC
metrics1_trv <- metrics1_trv_lidar$trv # 1ST TOTAL RECOVERABLE VOLUME METRIC
metrics1_saw <- metrics1_saw_lidar$saw # 1ST SAW METRIC
metrics1_industrial <- metrics1_industrial_lidar$industrial # 1ST INDUSTRIAL METRIC
metrics1_chip <- metrics1_chip_lidar$chip # 1ST CHIP METRIC
metrics1_pulp <- metrics1_pulp_lidar$pulp # 1ST PULP METRIC

metrics2_lidar <- data.frame(cbind(metrics2_file[, 11], metrics2_file[, 12:160])) # COMBINE AGE METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_lidar)[1] <- "age" # PROVIDE COLUMN NAME TO THE AGE METRIC FOR 2ND
metrics2_n_lidar <- data.frame(cbind(metrics2_file[, 2], metrics2_lidar)) # COMBINE STOCKING METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_n_lidar)[1] <- "n" # PROVIDE COLUMN NAME TO THE STOCKING METRIC FOR 2ND
metrics2_ba_lidar <- data.frame(cbind(metrics2_file[, 3], metrics2_lidar)) # COMBINE BASAL AREA METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_ba_lidar)[1] <- "ba" # PROVIDE COLUMN NAME TO THE BASAL AREA METRIC FOR 2ND
metrics2_h_lidar <- data.frame(cbind(metrics2_file[, 4], metrics2_lidar)) # COMBINE HEIGHT METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_h_lidar)[1] <- "h" # PROVIDE COLUMN NAME TO THE HEIGHT METRIC FOR 2ND
metrics2_vs_lidar <- data.frame(cbind(metrics2_file[, 5], metrics2_lidar)) # COMBINE STEM VOLUME METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_vs_lidar)[1] <- "vs" # PROVIDE COLUMN NAME TO THE STEM VOLUME METRIC FOR 2ND
metrics2_trv_lidar <- data.frame(cbind(metrics2_file[, 6], metrics2_lidar)) # COMBINE TOTAL RECOVERABLE VOLUME METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_trv_lidar)[1] <- "trv" # PROVIDE COLUMN NAME TO THE TOTAL RECOVERABLE VOLUME METRIC FOR 2ND
metrics2_saw_lidar <- data.frame(cbind(metrics2_file[, 7], metrics2_lidar))# COMBINE SAW METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_saw_lidar)[1] <- "saw" # PROVIDE COLUMN NAME TO THE SAW METRIC FOR 2ND
metrics2_industrial_lidar <- data.frame(cbind(metrics2_file[, 8], metrics2_lidar)) # COMBINE INDUSTRIAL METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_industrial_lidar)[1] <- "industrial" # PROVIDE COLUMN NAME TO THE INDUSTRIAL METRIC FOR 2ND
metrics2_chip_lidar <- data.frame(cbind(metrics2_file[, 9], metrics2_lidar)) # COMBINE CHIP METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_chip_lidar)[1] <- "chip" # PROVIDE COLUMN NAME TO THE CHIP METRIC FOR 2ND
metrics2_pulp_lidar <- data.frame(cbind(metrics2_file[, 10], metrics2_lidar)) # COMBINE PULP METRIC WITH THE LIDAR METRICS FOR 2ND
colnames(metrics2_pulp_lidar)[1] <- "pulp" # PROVIDE COLUMN NAME TO THE PULP METRIC FOR 2ND

metrics2_n <- metrics2_n_lidar$n  # 2ND STOCKING METRIC
metrics2_ba <- metrics2_ba_lidar$ba # 2ND BASAL AREA METRIC
metrics2_h <- metrics2_h_lidar$h # 2ND HEIGHT METRIC
metrics2_vs <- metrics2_vs_lidar$vs # 2ND STEM VOLUME METRIC
metrics2_trv <- metrics2_trv_lidar$trv # 2ND TOTAL RECOVERABLE VOLUME METRIC
metrics2_saw <- metrics2_saw_lidar$saw # 2ND SAW METRIC
metrics2_industrial <- metrics2_industrial_lidar$industrial # 2ND INDUSTRIAL METRIC
metrics2_chip <- metrics2_chip_lidar$chip # 2ND CHIP METRIC
metrics2_pulp <- metrics2_pulp_lidar$pulp # 2ND PULP METRIC

# DEVELOP THE SPATIAL-TEMPORAL RANDOMF FOREST MODEL AND IMPUTE THE GROUND INVENTORY METRICS
set.seed(1)
metrics1_n_rf <- randomForest(n~., data = metrics1_n_lidar, ntree = 100, mtry = floor(ncol(metrics1_n_lidar)/3),
                              importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM STOCKING BASED 1ST METRICS...
# ...TO IMPUTE 2ND STOCKING METRIC
varImpPlot(metrics1_n_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_n_imp_metrics1 <- predict(metrics1_n_rf, metrics2_lidar) #2ND STOCKING METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_ba_rf <- randomForest(ba~., data = metrics1_ba_lidar, ntree = 100, mtry = floor(ncol(metrics1_ba_lidar)/3),
                               importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM BASAL AREA BASED 1ST METRICS...
# ...TO IMPUTE 2ND BASAL AREA METRIC
varImpPlot(metrics1_ba_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_ba_imp_metrics1 <- predict(metrics1_ba_rf, metrics2_lidar) #2ND BASAL AREA METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_h_rf <- randomForest(h~., data = metrics1_h_lidar, ntree = 100, mtry = floor(ncol(metrics1_h_lidar)/3),
                              importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM HEIGHT BASED 1ST METRICS...
# ...TO IMPUTE 2ND HEIGHT METRIC
varImpPlot(metrics1_h_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_h_imp_metrics1 <- predict(metrics1_h_rf, metrics2_lidar) #2ND HEIGHT METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_vs_rf <- randomForest(vs~., data = metrics1_vs_lidar, ntree = 100, mtry = floor(ncol(metrics1_vs_lidar)/3),
                               importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM STEM VOLUME BASED 1ST METRICS...
# ...TO IMPUTE 2ND STEM VOLUME METRIC
varImpPlot(metrics1_vs_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_vs_imp_metrics1 <- predict(metrics1_vs_rf, metrics2_lidar) #2ND STEM VOLUME METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_trv_rf <- randomForest(trv~., data = metrics1_trv_lidar, ntree = 100, mtry = floor(ncol(metrics1_trv_lidar)/3),
                                importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM TOTAL RECOVERABLE VOLUME BASED 1ST METRICS...
# ...TO IMPUTE 2ND TOTAL RECOVERABLE VOLUME METRIC
varImpPlot(metrics1_trv_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_trv_imp_metrics1 <- predict(metrics1_trv_rf, metrics2_lidar) #2ND TOTAL RECOVERABLE VOLUME METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_saw_rf <- randomForest(saw~., data = metrics1_saw_lidar, ntree = 100, mtry = floor(ncol(metrics1_saw_lidar)/3),
                                importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM SAW BASED 1ST METRICS...
# ...TO IMPUTE 2ND SAW METRIC
varImpPlot(metrics1_saw_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_saw_imp_metrics1 <- predict(metrics1_saw_rf, metrics2_lidar) #2ND SAW METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_industrial_rf <- randomForest(industrial~., data = metrics1_industrial_lidar, ntree = 100, mtry = floor(ncol(metrics1_industrial_lidar)/3),
                                       importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM INDUSTRIAL BASED 1ST METRICS...
# ...TO IMPUTE 2ND INDUSTRIAL METRIC
varImpPlot(metrics1_industrial_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_industrial_imp_metrics1 <- predict(metrics1_industrial_rf, metrics2_lidar) #2ND INDUSTRIAL METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_chip_rf <- randomForest(chip~., data = metrics1_chip_lidar, ntree = 100, mtry = floor(ncol(metrics1_chip_lidar)/3),
                                 importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM CHIP BASED 1ST METRICS...
# ...TO IMPUTE 2ND CHIP METRIC
varImpPlot(metrics1_chip_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_chip_imp_metrics1 <- predict(metrics1_chip_rf, metrics2_lidar) #2ND CHIP METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL

metrics1_pulp_rf <- randomForest(pulp~., data = metrics1_pulp_lidar, ntree = 100, mtry = floor(ncol(metrics1_pulp_lidar)/3),
                                 importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM PULP BASED 1ST METRICS...
# ...TO IMPUTE 2ND PULP METRIC
varImpPlot(metrics1_pulp_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 1ST MODEL
metrics2_pulp_imp_metrics1 <- predict(metrics1_pulp_rf, metrics2_lidar) #2ND PULP METRIC IS IMPUTED FROM 1ST RANDOM FOREST MODEL


metrics2_n_rf <- randomForest(n~., data = metrics2_n_lidar, ntree = 100, mtry = floor(ncol(metrics2_n_lidar)/3),
                              importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM STOCKING BASED 2ND METRICS...
# ...TO IMPUTE 1ST STOCKING METRIC
varImpPlot(metrics2_n_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_n_imp_metrics2 <- predict(metrics2_n_rf, metrics1_lidar) #1ST STOCKING METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_ba_rf <- randomForest(ba~., data = metrics2_ba_lidar, ntree = 100, mtry = floor(ncol(metrics2_ba_lidar)/3),
                               importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM BASAL AREA BASED 2ND METRICS...
# ...TO IMPUTE 1ST BASAL AREA METRIC
varImpPlot(metrics2_ba_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_ba_imp_metrics2 <- predict(metrics2_ba_rf, metrics1_lidar) #1ST BASAL AREA METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_h_rf <- randomForest(h~., data = metrics2_h_lidar, ntree = 100, mtry = floor(ncol(metrics2_h_lidar)/3),
                              importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM HEIGHT BASED 2ND METRICS...
# ...TO IMPUTE 1ST HEIGHT METRIC
varImpPlot(metrics2_h_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_h_imp_metrics2 <- predict(metrics2_h_rf, metrics1_lidar) #1ST HEIGHT METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_vs_rf <- randomForest(vs~., data = metrics2_vs_lidar, ntree = 100, mtry = floor(ncol(metrics2_vs_lidar)/3),
                               importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM STEM VOLUME BASED 2ND METRICS...
# ...TO IMPUTE 1ST STEM VOLUME METRIC
varImpPlot(metrics2_vs_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_vs_imp_metrics2 <- predict(metrics2_vs_rf, metrics1_lidar) #1ST STEM VOLUME METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_trv_rf <- randomForest(trv~., data = metrics2_trv_lidar, ntree = 100, mtry = floor(ncol(metrics2_trv_lidar)/3),
                                importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM TOTAL RECOVERABLE VOLUME BASED 2ND METRICS...
# ...TO IMPUTE 1ST TOTAL RECOVERABLE VOLUME METRIC
varImpPlot(metrics2_trv_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_trv_imp_metrics2 <- predict(metrics2_trv_rf, metrics1_lidar) #1ST TOTAL RECOVERABLE VOLUME METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_saw_rf <- randomForest(saw~., data = metrics2_saw_lidar, ntree = 100, mtry = floor(ncol(metrics2_saw_lidar)/3),
                                importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM SAW BASED 2ND METRICS...
# ...TO IMPUTE 1ST SAW METRIC
varImpPlot(metrics2_saw_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_saw_imp_metrics2 <- predict(metrics2_saw_rf, metrics1_lidar) #1ST SAW METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_industrial_rf <- randomForest(industrial~., data = metrics2_industrial_lidar, ntree = 100, mtry = floor(ncol(metrics2_industrial_lidar)/3),
                                       importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM INDUSTRIAL BASED 2ND METRICS...
# ...TO IMPUTE 1ST INDUSTRIAL METRIC
varImpPlot(metrics2_industrial_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_industrial_imp_metrics2 <- predict(metrics2_industrial_rf, metrics1_lidar) #1ST INDUSTRIAL METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_chip_rf <- randomForest(chip~., data = metrics2_chip_lidar, ntree = 100, mtry = floor(ncol(metrics2_chip_lidar)/3),
                                 importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM CHIP BASED 2ND METRICS...
# ...TO IMPUTE 1ST CHIP METRIC
varImpPlot(metrics2_chip_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_chip_imp_metrics2 <- predict(metrics2_chip_rf, metrics1_lidar) #1ST CHIP METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

metrics2_pulp_rf <- randomForest(pulp~., data = metrics2_pulp_lidar, ntree = 100, mtry = floor(ncol(metrics2_pulp_lidar)/3),
                                 importance = TRUE, na.action = na.omit) # RANDOM FOREST MODEL IS DEVELOPED FROM PULP BASED 2ND METRICS...
# ...TO IMPUTE 1ST PULP METRIC
varImpPlot(metrics2_pulp_rf, sort = TRUE) # PLOT VARIABLE IMPORTANCE WITHIN 2ND MODEL
metrics1_pulp_imp_metrics2 <- predict(metrics2_pulp_rf, metrics1_lidar) #1ST PULP METRIC IS IMPUTED FROM 2ND RANDOM FOREST MODEL

# ACCURACY ASSESSMENT OF IMPUTED 1ST AND 2ND GROUND INVENTORY METRICS WITH RESPECT TO THEIR ORIGINAL METRIC VALUES...
# ...BASED ON THE ESTIMATION OF NORMALIZED RMSE (NRMSE) AND PERCENT BIASNESS (PBIAS) MEASURES
set.seed(2)
metrics1_n_nrmse_metrics2 <- round(nrmse(metrics1_n_imp_metrics2, metrics1_n, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_ba_nrmse_metrics2 <- round(nrmse(metrics1_ba_imp_metrics2, metrics1_ba, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_h_nrmse_metrics2 <- round(nrmse(metrics1_h_imp_metrics2, metrics1_h, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_vs_nrmse_metrics2 <- round(nrmse(metrics1_vs_imp_metrics2, metrics1_vs, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_trv_nrmse_metrics2 <- round(nrmse(metrics1_trv_imp_metrics2, metrics1_trv, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_saw_nrmse_metrics2 <- round(nrmse(metrics1_saw_imp_metrics2, metrics1_saw, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_industrial_nrmse_metrics2 <- round(nrmse(metrics1_industrial_imp_metrics2, metrics1_industrial, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_chip_nrmse_metrics2 <- round(nrmse(metrics1_chip_imp_metrics2, metrics1_chip, na.rm = TRUE, norm = "maxmin"), 3)
metrics1_pulp_nrmse_metrics2 <- round(nrmse(metrics1_pulp_imp_metrics2, metrics1_pulp, na.rm = TRUE, norm = "maxmin"), 3)

metrics1_n_pbias_metrics2 <- round(pbias(metrics1_n_imp_metrics2, metrics1_n, na.rm = TRUE), 3)
metrics1_ba_pbias_metrics2 <- round(pbias(metrics1_ba_imp_metrics2, metrics1_ba, na.rm = TRUE), 3)
metrics1_h_pbias_metrics2 <- round(pbias(metrics1_h_imp_metrics2, metrics1_h, na.rm = TRUE), 3)
metrics1_vs_pbias_metrics2 <- round(pbias(metrics1_vs_imp_metrics2, metrics1_vs, na.rm = TRUE), 3)
metrics1_trv_pbias_metrics2 <- round(pbias(metrics1_trv_imp_metrics2, metrics1_trv, na.rm = TRUE), 3)
metrics1_saw_pbias_metrics2 <- round(pbias(metrics1_saw_imp_metrics2, metrics1_saw, na.rm = TRUE), 3)
metrics1_industrial_pbias_metrics2 <- round(pbias(metrics1_industrial_imp_metrics2, metrics1_industrial, na.rm = TRUE), 3)
metrics1_chip_pbias_metrics2 <- round(pbias(metrics1_chip_imp_metrics2, metrics1_chip, na.rm = TRUE), 3)
metrics1_pulp_pbias_metrics2 <- round(pbias(metrics1_pulp_imp_metrics2, metrics1_pulp, na.rm = TRUE), 3)

metrics2_n_nrmse_metrics1 <- round(nrmse(metrics2_n_imp_metrics1, metrics2_n, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_ba_nrmse_metrics1 <- round(nrmse(metrics2_ba_imp_metrics1, metrics2_ba, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_h_nrmse_metrics1 <- round(nrmse(metrics2_h_imp_metrics1, metrics2_h, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_vs_nrmse_metrics1 <- round(nrmse(metrics2_vs_imp_metrics1, metrics2_vs, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_trv_nrmse_metrics1 <- round(nrmse(metrics2_trv_imp_metrics1, metrics2_trv, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_saw_nrmse_metrics1 <- round(nrmse(metrics2_saw_imp_metrics1, metrics2_saw, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_industrial_nrmse_metrics1 <- round(nrmse(metrics2_industrial_imp_metrics1, metrics2_industrial, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_chip_nrmse_metrics1 <- round(nrmse(metrics2_chip_imp_metrics1, metrics2_chip, na.rm = TRUE, norm = "maxmin"), 3)
metrics2_pulp_nrmse_metrics1 <- round(nrmse(metrics2_pulp_imp_metrics1, metrics2_pulp, na.rm = TRUE, norm = "maxmin"), 3)

metrics2_n_pbias_metrics1 <- round(pbias(metrics2_n_imp_metrics1, metrics2_n, na.rm = TRUE), 3)
metrics2_ba_pbias_metrics1 <- round(pbias(metrics2_ba_imp_metrics1, metrics2_ba, na.rm = TRUE), 3)
metrics2_h_pbias_metrics1 <- round(pbias(metrics2_h_imp_metrics1, metrics2_h, na.rm = TRUE), 3)
metrics2_vs_pbias_metrics1 <- round(pbias(metrics2_vs_imp_metrics1, metrics2_vs, na.rm = TRUE), 3)
metrics2_trv_pbias_metrics1 <- round(pbias(metrics2_trv_imp_metrics1, metrics2_trv, na.rm = TRUE), 3)
metrics2_saw_pbias_metrics1 <- round(pbias(metrics2_saw_imp_metrics1, metrics2_saw, na.rm = TRUE), 3)
metrics2_industrial_pbias_metrics1 <- round(pbias(metrics2_industrial_imp_metrics1, metrics2_industrial, na.rm = TRUE), 3)
metrics2_chip_pbias_metrics1 <- round(pbias(metrics2_chip_imp_metrics1, metrics2_chip, na.rm = TRUE), 3)
metrics2_pulp_pbias_metrics1 <- round(pbias(metrics2_pulp_imp_metrics1, metrics2_pulp, na.rm = TRUE), 3)

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

metrics1_GI_imp_metrics2 <- cbind(metrics1_n_imp_metrics2, metrics1_ba_imp_metrics2, metrics1_h_imp_metrics2,
                                  metrics1_vs_imp_metrics2, metrics1_trv_imp_metrics2, metrics1_saw_imp_metrics2,
                                  metrics1_industrial_imp_metrics2, metrics1_chip_imp_metrics2, metrics1_pulp_imp_metrics2,
                                  metrics1_n, metrics1_ba, metrics1_h,
                                  metrics1_vs, metrics1_trv, metrics1_saw,
                                  metrics1_industrial, metrics1_chip, metrics1_pulp)

colnames(metrics1_GI_imp_metrics2)[1] <- "n"
colnames(metrics1_GI_imp_metrics2)[2] <- "ba"
colnames(metrics1_GI_imp_metrics2)[3] <- "h"
colnames(metrics1_GI_imp_metrics2)[4] <- "vs"
colnames(metrics1_GI_imp_metrics2)[5] <- "trv"
colnames(metrics1_GI_imp_metrics2)[6] <- "saw"
colnames(metrics1_GI_imp_metrics2)[7] <- "industrial"
colnames(metrics1_GI_imp_metrics2)[8] <- "chip"
colnames(metrics1_GI_imp_metrics2)[9] <- "pulp"
colnames(metrics1_GI_imp_metrics2)[10] <- "n.o"
colnames(metrics1_GI_imp_metrics2)[11] <- "ba.o"
colnames(metrics1_GI_imp_metrics2)[12] <- "h.o"
colnames(metrics1_GI_imp_metrics2)[13] <- "vs.o"
colnames(metrics1_GI_imp_metrics2)[14] <- "trv.o"
colnames(metrics1_GI_imp_metrics2)[15] <- "saw.o"
colnames(metrics1_GI_imp_metrics2)[16] <- "industrial.o"
colnames(metrics1_GI_imp_metrics2)[17] <- "chip.o"
colnames(metrics1_GI_imp_metrics2)[18] <- "pulp.o"
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

metrics2_GI_imp_metrics1 <- cbind(metrics2_n_imp_metrics1, metrics2_ba_imp_metrics1, metrics2_h_imp_metrics1,
                                     metrics2_vs_imp_metrics1, metrics2_trv_imp_metrics1, metrics2_saw_imp_metrics1,
                                     metrics2_industrial_imp_metrics1, metrics2_chip_imp_metrics1, metrics2_pulp_imp_metrics1,
                                     metrics2_n, metrics2_ba, metrics2_h,
                                     metrics2_vs, metrics2_trv, metrics2_saw,
                                     metrics2_industrial, metrics2_chip, metrics2_pulp)

colnames(metrics2_GI_imp_metrics1)[1] <- "n"
colnames(metrics2_GI_imp_metrics1)[2] <- "ba"
colnames(metrics2_GI_imp_metrics1)[3] <- "h"
colnames(metrics2_GI_imp_metrics1)[4] <- "vs"
colnames(metrics2_GI_imp_metrics1)[5] <- "trv"
colnames(metrics2_GI_imp_metrics1)[6] <- "saw"
colnames(metrics2_GI_imp_metrics1)[7] <- "industrial"
colnames(metrics2_GI_imp_metrics1)[8] <- "chip"
colnames(metrics2_GI_imp_metrics1)[9] <- "pulp"
colnames(metrics2_GI_imp_metrics1)[10] <- "n.o"
colnames(metrics2_GI_imp_metrics1)[11] <- "ba.o"
colnames(metrics2_GI_imp_metrics1)[12] <- "h.o"
colnames(metrics2_GI_imp_metrics1)[13] <- "vs.o"
colnames(metrics2_GI_imp_metrics1)[14] <- "trv.o"
colnames(metrics2_GI_imp_metrics1)[15] <- "saw.o"
colnames(metrics2_GI_imp_metrics1)[16] <- "industrial.o"
colnames(metrics2_GI_imp_metrics1)[17] <- "chip.o"
colnames(metrics2_GI_imp_metrics1)[18] <- "pulp.o"

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










