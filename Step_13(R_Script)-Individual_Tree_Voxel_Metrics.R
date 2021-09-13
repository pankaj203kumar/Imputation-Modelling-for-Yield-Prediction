# THIS CODE ESTIMATES THE INDIVIDUAL TREE BASED VOXEL METRICS FROM HEIGHT NORMALIZED LIDAR POINT CLOUD DATA
# -----------------------------------------------

rm(list = ls()) #REMOVES ALL THE VARIABLES
cat("\f") #CLEAR SCREEN

setwd("SET WORKING DIRECTORY")

require(openxlsx)
require(matlib)
require(Metrics)
require(ggplot2)
require(lidR)
require(plot3D)
require(rgl)
require(rlas)
require(sf)

# READ THE LIDAR FILE IN LAZ FORMAT

start.time = Sys.time()

lidar_input_file_path = "SET INPUT FOLDER WHERE INDIVIDUAL TREE BASED HEIGHT NORMALIZED LIDAR FILES ARE LOCATED"

lidar_file_list = list.files(path = lidar_input_file_path, pattern = "\\.laz$", full.names = TRUE, include.dirs = TRUE) # LISTS ALL THE LIDAR FILE NAMES
lidar_file_list_plotid = list.files(path = lidar_input_file_path, pattern = "\\.laz$") # LISTS PLOT IDS ASSOCIATED WITH THE LIDAR FILES

voxel_metrics_plots = data.frame()
pb <- winProgressBar(title = "Progress Bar", min = 1, max = length(lidar_file_list), initial = 1, width = 300)
for (m in 1:length(lidar_file_list))
{
  lidar_file_unfilt = readLAS(lidar_file_list[m], select = "xyzi")
  
  lidar_file_plotid = as.numeric(substr(lidar_file_list_plotid[m], 1, nchar(lidar_file_list_plotid[m]) - 4))
  
  lidar_file = readLAS(lidar_file_list[m], select = "xyzi", filter = "-drop_z_below 0.01")
  lidar_x = cloud_metrics(lidar_file, ~c(X))
  lidar_y = cloud_metrics(lidar_file, ~c(Y))
  lidar_z = cloud_metrics(lidar_file, ~c(Z))
  lidar_i = cloud_metrics(lidar_file, ~c(Intensity))
  
  lidar_x_range = max(lidar_x) - min(lidar_x)
  lidar_y_range = max(lidar_y) - min(lidar_y)
  lidar_z_range = max(lidar_z) - min(lidar_z)
  
  # GENERATE VERTICAL COMPLEXITY INDEX (VCI) METRICS FROM 0.3m,0.5m,1m,1.5m,2m,2.5m,3m XYZ VOXEL RESOLUTION
  
  vci_0.3 = VCI(lidar_z, zmax = max(lidar_z), by = 0.3)
  vci_0.5 = VCI(lidar_z, zmax = max(lidar_z), by = 0.5)
  vci_1 = VCI(lidar_z, zmax = max(lidar_z), by = 1)
  vci_1.5 = VCI(lidar_z, zmax = max(lidar_z), by = 1.5)
  vci_2 = VCI(lidar_z, zmax = max(lidar_z), by = 2)
  vci_2.5 = VCI(lidar_z, zmax = max(lidar_z), by = 2.5)
  vci_3 = VCI(lidar_z, zmax = max(lidar_z), by = 3)
  
  # GENERATE COEFFICIENT OF VARIATION FOR LEAF AREA DENSITY (CVLAD) METRIC FROM 0.3m,0.5m,1m,1.5m,2m,2.5m,3m XYZ VOXEL RESOLUTION
  
  lad_0.3 = LAD(lidar_z, 0.3, k = 0.5, z0 = 2)
  normlad_0.3 = (lad_0.3$lad)/0.3
  cvlad_0.3 = sd(normlad_0.3)/mean(normlad_0.3)
  
  lad_0.5 = LAD(lidar_z, 0.5, k = 0.5, z0 = 2)
  normlad_0.5 = (lad_0.5$lad)/0.5
  cvlad_0.5 = sd(normlad_0.5)/mean(normlad_0.5)
  
  lad_1 = LAD(lidar_z, 1, k = 0.5, z0 = 2)
  normlad_1 = (lad_1$lad)/1
  cvlad_1 = sd(normlad_1)/mean(normlad_1)
  
  lad_1.5 = LAD(lidar_z, 1.5, k = 0.5, z0 = 2)
  normlad_1.5 = (lad_1.5$lad)/1.5
  cvlad_1.5 = sd(normlad_1.5)/mean(normlad_1.5)
  
  lad_2 = LAD(lidar_z, 2, k = 0.5, z0 = 2)
  normlad_2 = (lad_2$lad)/2
  cvlad_2 = sd(normlad_2)/mean(normlad_2)
  
  lad_2.5 = LAD(lidar_z, 2.5, k = 0.5, z0 = 2)
  normlad_2.5 = (lad_2.5$lad)/2.5
  cvlad_2.5 = sd(normlad_2.5)/mean(normlad_2.5)
  
  lad_3 = LAD(lidar_z, 3, k = 0.5, z0 = 2)
  normlad_3 = (lad_3$lad)/3
  cvlad_3 = sd(normlad_3)/mean(normlad_3)
  
  # DEVELOP VOXELS TO GENERATE UNIVARIATE BIOMASS BASED VOXEL METRICS
  
  sv_attributes = function(z, i)
  {
    ret = list(
      no_points = as.integer(length(z)),
      median_z = as.double(median(z)),
      median_i = as.integer(median(i))
    )
    return(ret)
  }
  
  # GENERATE BIOMASS BASED SUB-VOXEL (SV) METRIC FROM 5m,10m,15m,20m,25m,30m,40m,45m Z VOXEL RESOLUTION. THE METRICS GENERATED ARE BASED...
  # ...ON NUMBER OF POINTS (P_SV), FREQUENCY RATIO (FR_SV), MEDIAN INTENSITY (Imed_SV)
  
  sv_voxel = voxel_metrics(lidar_file, ~sv_attributes(Z, Intensity), res = c(1000000, 5)) # RES REPRESENTS THE VOXEL SIZE. THE XY VOXEL SIZE IS KEPT...
  # ...HIGH TO GENERATE VOXEL METRICS. THE Z VOXEL SIZE OF 5m WILL PROVIDE THE METRICS OF POINTS IN BETWEEN 2.5m & 7.5mN HEIGHT, SIMILARLY...
  # ...10m z VOXEL SIZE WILL PROVIDE THE METRICS IN BETWEEN 7.5m & 12.5m HEIGHT.
  
  sv_p5 = 0
  sv_p10 = 0
  sv_p15 = 0
  sv_p20 = 0
  sv_p25 = 0
  sv_p30 = 0
  sv_p35 = 0
  sv_p40 = 0
  sv_p45 = 0
  sv_imed5 = 0
  sv_imed10 = 0
  sv_imed15 = 0
  sv_imed20 = 0
  sv_imed25 = 0
  sv_imed30 = 0
  sv_imed35 = 0
  sv_imed40 = 0
  sv_imed45 = 0
  sv_p = 0
  for (n in 1:dim(sv_voxel)[1])
  {
    if (sv_voxel[n]$Z == 5)
    {
      sv_p5 = sv_p5 + sv_voxel[n]$no_points
      sv_imed5 = sv_imed5 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 10)
    {
      sv_p10 = sv_p10 + sv_voxel[n]$no_points
      sv_imed10 = sv_imed10 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 15)
    {
      sv_p15 = sv_p15 + sv_voxel[n]$no_points
      sv_imed15 = sv_imed15 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 20)
    {
      sv_p20 = sv_p20 + sv_voxel[n]$no_points
      sv_imed20 = sv_imed20 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 25)
    {
      sv_p25 = sv_p25 + sv_voxel[n]$no_points
      sv_imed25 = sv_imed25 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 30)
    {
      sv_p30 = sv_p30 + sv_voxel[n]$no_points
      sv_imed30 = sv_imed30 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 35)
    {
      sv_p35 = sv_p35 + sv_voxel[n]$no_points
      sv_imed35 = sv_imed35 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 40)
    {
      sv_p40 = sv_p40 + sv_voxel[n]$no_points
      sv_imed40 = sv_imed40 + sv_voxel[n]$median_i
    }
    if (sv_voxel[n]$Z == 45)
    {
      sv_p45 = sv_p45 + sv_voxel[n]$no_points
      sv_imed45 = sv_imed45 + sv_voxel[n]$median_i
    }
    sv_p = sv_p + sv_voxel[n]$no_points
  }
  
  sv_fr5 = sv_p5/sv_p
  sv_fr10 = sv_p10/sv_p
  sv_fr15 = sv_p15/sv_p
  sv_fr20 = sv_p20/sv_p
  sv_fr25 = sv_p25/sv_p
  sv_fr30 = sv_p30/sv_p
  sv_fr35 = sv_p35/sv_p
  sv_fr40 = sv_p40/sv_p
  sv_fr45 = sv_p45/sv_p
  
  # GENERATE BIOMASS BASED DENSITY (D) METRIC FROM 5m,10m,15m,20m,25m,30m,40m,45m Z VOXEL RESOLUTION. THE METRICS GENERATED ARE BASED ON...
  # ...NUMBER OF POINTS (P_D), FREQUENCY RATIO (FR_D), MEDIAN INTENSITY (Imed_D)
  
  d_p5 = 0
  d_p10 = 0
  d_p15 = 0
  d_p20 = 0
  d_p25 = 0
  d_p30 = 0
  d_p35 = 0
  d_p40 = 0
  d_p45 = 0
  d_imed5 = 0
  d_imed10 = 0
  d_imed15 = 0
  d_imed20 = 0
  d_imed25 = 0
  d_imed30 = 0
  d_imed35 = 0
  d_imed40 = 0
  d_imed45 = 0
  d_p = 0
  for (p in 1:dim(sv_voxel)[1])
  {
    if (sv_voxel[p]$Z >= 5)
    {
      d_p5 = d_p5 + sv_voxel[p]$no_points
      d_imed5 = d_imed5 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 10)
    {
      d_p10 = d_p10 + sv_voxel[p]$no_points
      d_imed10 = d_imed10 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 15)
    {
      d_p15 = d_p15 + sv_voxel[p]$no_points
      d_imed15 = d_imed15 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 20)
    {
      d_p20 = d_p20 + sv_voxel[p]$no_points
      d_imed20 = d_imed20 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 25)
    {
      d_p25 = d_p25 + sv_voxel[p]$no_points
      d_imed25 = d_imed25 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 30)
    {
      d_p30 = d_p30 + sv_voxel[p]$no_points
      d_imed30 = d_imed30 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 35)
    {
      d_p35 = d_p35 + sv_voxel[p]$no_points
      d_imed35 = d_imed35 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 40)
    {
      d_p40 = d_p40 + sv_voxel[p]$no_points
      d_imed40 = d_imed40 + sv_voxel[p]$median_i
    }
    if (sv_voxel[p]$Z >= 45)
    {
      d_p45 = d_p45 + sv_voxel[p]$no_points
      d_imed45 = d_imed45 + sv_voxel[p]$median_i
    }
    d_p = d_p + sv_voxel[p]$no_points
  }
  
  d_fr5 = d_p5/d_p
  d_fr10 = d_p10/d_p
  d_fr15 = d_p15/d_p
  d_fr20 = d_p20/d_p
  d_fr25 = d_p25/d_p
  d_fr30 = d_p30/d_p
  d_fr35 = d_p35/d_p
  d_fr40 = d_p40/d_p
  d_fr45 = d_p45/d_p
  
  # GENERATE BIOMASS BASED SUB-VOXEL MAXIMUM (SVM) METRIC FROM 5m,10m,15m,20m,25m,30m,40m,45m Z VOXEL RESOLUTION. THE METRICS GENERATED ARE BASED ON...
  # ...MEDIAN HEIGHT (Hmed_SVM) & FREQUENCY (F_SVM)
  
  svm_voxel = data.frame()
  for (r in 1:dim(sv_voxel)[1])
  {
    if (sv_voxel[r]$Z >= 5)
    {
      svm_voxel = rbind(svm_voxel, sv_voxel[r])
    }
  }
  
  svm_maxp_index = which.max(svm_voxel$no_points)
  svm_maxp_hmed = svm_voxel[svm_maxp_index]$median_z
  
  svm_voxel_maxp_above = data.frame()
  svm_voxel_maxp_above = rbind(svm_voxel_maxp_above, svm_voxel[svm_maxp_index])
  for (s in 1:dim(svm_voxel)[1])
  {
    if ((svm_voxel[s]$Z > svm_voxel[svm_maxp_index]$Z))
    {
      svm_voxel_maxp_above = rbind(svm_voxel_maxp_above, svm_voxel[s])
    }
  }
  
  svm_maxp_above_p = 0
  svm_maxp_above_f = 0
  if (dim(svm_voxel_maxp_above)[1] > 1)
  {
    for (t in 2:dim(svm_voxel_maxp_above)[1])
    {
      svm_maxp_above_p = svm_maxp_above_p + svm_voxel_maxp_above[t]$no_points
    }
    svm_maxp_above_f = svm_maxp_above_p/svm_voxel_maxp_above[1]$no_points
  }
  
  # GENERATE CANOPY CLOSURE (CC_above) AND MEAN PERCENTAGE CANOPY CLOSURE (per_cc_above) AT HEIGHT ABOVE 5m,10m,15m,20m,25m FROM...
  # ...1m XYZ VOXEL RESOLUTION
  
  cc_attributes = function(z)
  {
    ret = list(
      no_points = as.integer(length(z))
    )
    return(ret)
  }
  
  cc_voxel = voxel_metrics(lidar_file, ~cc_attributes(Z), res = c(1, 1)) # RES REPRESENTS THE SIZE OF VOXEL
  
  cc_above5_voxel = data.frame()
  cc_above5_p = 0
  cc_p = 0
  for (u in 1:dim(cc_voxel)[1])
  {
    cc_p = cc_p + cc_voxel[u]$no_points
  }
  
  for (u in 1:dim(cc_voxel)[1])
  {
    if (cc_voxel[u]$Z >= 5)
    {
      cc_above5_voxel = rbind(cc_above5_voxel, cc_voxel[u])
      cc_above5_p = cc_above5_p + cc_voxel[u]$no_points
    }
  }
  
  cc_above5 = dim(cc_above5_voxel)[1]/dim(cc_voxel)[1]
  per_cc_above5 = (cc_above5_p / cc_p) * 100
  
  cc_above10_voxel = data.frame()
  cc_above10_p = 0
  for (u in 1:dim(cc_voxel)[1])
  {
    if (cc_voxel[u]$Z >= 10)
    {
      cc_above10_voxel = rbind(cc_above10_voxel, cc_voxel[u])
      cc_above10_p = cc_above10_p + cc_voxel[u]$no_points
    }
  }
  
  cc_above10 = dim(cc_above10_voxel)[1]/dim(cc_voxel)[1]
  per_cc_above10 = (cc_above10_p / cc_p) * 100
  
  cc_above15_voxel = data.frame()
  cc_above15_p = 0
  for (u in 1:dim(cc_voxel)[1])
  {
    if (cc_voxel[u]$Z >= 15)
    {
      cc_above15_voxel = rbind(cc_above15_voxel, cc_voxel[u])
      cc_above15_p = cc_above15_p + cc_voxel[u]$no_points
    }
  }
  
  cc_above15 = dim(cc_above15_voxel)[1]/dim(cc_voxel)[1]
  per_cc_above15 = (cc_above15_p / cc_p) * 100
  
  cc_above20_voxel = data.frame()
  cc_above20_p = 0
  for (u in 1:dim(cc_voxel)[1])
  {
    if (cc_voxel[u]$Z >= 20)
    {
      cc_above20_voxel = rbind(cc_above20_voxel, cc_voxel[u])
      cc_above20_p = cc_above20_p + cc_voxel[u]$no_points
    }
  }
  
  cc_above20 = dim(cc_above20_voxel)[1]/dim(cc_voxel)[1]
  per_cc_above20 = (cc_above20_p / cc_p) * 100
  
  cc_above25_voxel = data.frame()
  cc_above25_p = 0
  for (u in 1:dim(cc_voxel)[1])
  {
    if (cc_voxel[u]$Z >= 25)
    {
      cc_above25_voxel = rbind(cc_above25_voxel, cc_voxel[u])
      cc_above25_p = cc_above25_p + cc_voxel[u]$no_points
    }
  }
  
  cc_above25 = dim(cc_above25_voxel)[1]/dim(cc_voxel)[1]
  per_cc_above25 = (cc_above25_p / cc_p) * 100
  
  # GENERATE EFFECTIVE NUMBER OF LAYERS (d0_enl, d1_enl, d2_enl) FROM 1m XY AND 0.5m Z VOXEL RESOLUTION
  
  enl_attributes = function(z)
  {
    ret = list(
      no_points = as.integer(length(z))
    )
    return(ret)
  }
  
  enl_voxel = voxel_metrics(lidar_file, ~enl_attributes(Z), res = c(1, 0.5)) # RES REPRESENTS THE SIZE OF VOXEL
  
  enl_d0 = 0
  enl_d1_lg = 0
  enl_d2_sq = 0
  enl_max_z = max(enl_voxel$Z)
  for (x in seq(0.5, enl_max_z, 0.5))
  {
    lay_nvox = 0
    for (y in 1:dim(enl_voxel)[1])
    {
      if (enl_voxel[y]$Z == x)
      {
        lay_nvox = lay_nvox + 1
      }
    }
    lay_meanvox = lay_nvox/dim(enl_voxel)[1]
    enl_d0 = enl_d0 + lay_meanvox
    enl_d1_lg_pre = (lay_meanvox * log(lay_meanvox))
    if (!is.na(enl_d1_lg_pre))
    {
      enl_d1_lg = enl_d1_lg + enl_d1_lg_pre
    }
    enl_d2_sq = enl_d2_sq + (lay_meanvox^2)
  }
  enl_d1 = exp(-enl_d1_lg)
  enl_d2 = 1/enl_d2_sq
  
  # COMBINE ALL THE GENERATED VOXEL METRICS
  
  voxel_metrics_plot = round(cbind(lidar_file_plotid,
                                   vci_0.3, vci_0.5, vci_1, vci_1.5, vci_2, vci_2.5, vci_3,
                                   cvlad_0.3, cvlad_0.5, cvlad_1, cvlad_1.5, cvlad_2, cvlad_2.5, cvlad_3,
                                   sv_p5, sv_p10, sv_p15, sv_p20, sv_p25, sv_p30, sv_p35, sv_p40, sv_p45,
                                   sv_imed5, sv_imed10, sv_imed15, sv_imed20, sv_imed25, sv_imed30, sv_imed35, sv_imed40, sv_imed45,
                                   sv_fr5, sv_fr10, sv_fr15, sv_fr20, sv_fr25, sv_fr30, sv_fr35, sv_fr40, sv_fr45,
                                   d_p5, d_p10, d_p15, d_p20, d_p25, d_p30, d_p35, d_p40, d_p45,
                                   d_imed5, d_imed10, d_imed15, d_imed20, d_imed25, d_imed30, d_imed35, d_imed40, d_imed45,
                                   d_fr5, d_fr10, d_fr15, d_fr20, d_fr25, d_fr30, d_fr35, d_fr40, d_fr45,
                                   svm_maxp_hmed, svm_maxp_above_f,
                                   cc_above5, cc_above10, cc_above15, cc_above20, cc_above25,
                                   per_cc_above5, per_cc_above10, per_cc_above15, per_cc_above20, per_cc_above25,
                                   enl_d0, enl_d1, enl_d2), 3)
  
  voxel_metrics_plots = rbind(voxel_metrics_plots, voxel_metrics_plot)
  Sys.sleep(0.1)
  setWinProgressBar(pb, m, title = paste(round((m * 100)/length(lidar_file_list), 0), "% done"))
}
close(pb)

# WRITE OUTPUT INDIVIDUAL TREE BASED VOXEL METRICS FILE IN .XLSX FORMAT

voxel_metrics_file <- createWorkbook("Voxel_Metrics")
addWorksheet(voxel_metrics_file, "Voxel_Metrics")
writeData(voxel_metrics_file, sheet = "Voxel_Metrics", voxel_metrics_plots)
saveWorkbook(voxel_metrics_file,
             "WRITE FOLDER PATH AND OUTPUT INDIVIDUAL TREE BASED VOXEL METRICS FILE IN .XLSX FORMAT",
             overwrite = TRUE)

end.time = Sys.time()

time.taken = end.time - start.time
print(time.taken)
