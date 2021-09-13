# THIS CODE INTEGRATES PLOT BASED GROUND INVENTORY METRICS WITH THE LIDAR DERIVED STANDARD AND VOXEL METRICS BASED ON PLOT_IDs
# -----------------------------------------------

rm(list = ls()) #REMOVES ALL THE VARIABLES
cat("\f") #CLEAR SCREEN

setwd("SET WORKING DIRECTORY")

require(openxlsx)
require(matlib)
require(Metrics)
require(dplyr)

start.time = Sys.time()

# FIRST, INTEGRATES PLOT BASED GROUND INVENTORY METRICS WITH LIDAR DERIVED STANDARD METRICS

GI_file = read.xlsx("SET INPUT GROUND INVENTORY METRICS FILE IN .XLSX FORMAT") #PLOT_ID,X,Y AND GROUND INVENTORY METRICS
standard_file = read.xlsx("SET INPUT LIDAR DERIVED STANDARD METRICS FILE IN .XLSX FORMAT") #PLOT_ID AND LIDAR DERIVED STANDARD METRICS

View(GI_file)
View(standard_file)

GI = data.frame(GI_file)
standard = data.frame(standard_file)

GI_standard = data.frame()

pb <- winProgressBar(title = "Progress Bar", min = 1, max = nrow(GI), initial = 1, width = 300)
for (m in 1:nrow(GI))
{
  for (n in 1:nrow(standard))
  {
    if (GI[m, 1] == standard[n, 2]) # MATCHES PLOT_IDs FOR BOTH GROUND INVENTORY AND STANDARD FILES
    {
      GI_standard_col = cbind(GI[m, 1], GI[m, 4:13], standard[n, 9:74])
      GI_standard = rbind(GI_standard, GI_standard_col)
    }
  }
  Sys.sleep(0.1)
  setWinProgressBar(pb, m, title = paste(round((m * 100)/nrow(GI), 0), "% done"))
}
close(pb)

GI_standard_file <- createWorkbook("GI_standard_metrics")
addWorksheet(GI_standard_file, "GI_standard_metrics")
writeData(GI_standard_file, sheet = "GI_standard_metrics", GI_standard)
saveWorkbook(GI_standard_file,
             "WRITE FOLDER PATH AND OUTPUT INTEGRATED PLOT BASED GROUND INVENTORY-STANDARD METRICS FILE IN .XLSX FORMAT",
             overwrite = TRUE)

end.time = Sys.time()
time.taken = end.time - start.time
print(time.taken)

# SECOND, INTEGRATES PLOT BASED GROUND INVENTORY-STANDARD METRICS WITH LIDAR DERIVED VOXEL METRICS

voxel_file = read.xlsx("SET INPUT VOXEL METRICS FILE") #THE LIDAR DERIVED VOXEL METRIC FILE IN .XLSX FORMAT...
# ...WITH PLOT_ID AND DERIVED METRICS
View(voxel_file)

voxel = data.frame(voxel_file)

GI_standard_voxel = data.frame()

pb <- winProgressBar(title = "Progress Bar", min = 1, max = nrow(GI_standard), initial = 1, width = 300)
for (m in 1:nrow(GI_standard))
{
  for (n in 1:nrow(voxel))
  {
    if (GI_standard[m, 1] == voxel[n, 1]) # MATCHES PLOT_IDs FOR BOTH GROUND INVENTORY-STANDARD AND VOXEL FILES
    {
      GI_standard_voxel_col = cbind(GI_standard[m, 1:2], GI_standard[m, 3:78], voxel[n, 2:84])
      GI_standard_voxel = rbind(GI_standard_voxel, GI_standard_voxel_col)
    }
  }
  Sys.sleep(0.1)
  setWinProgressBar(pb, m, title = paste(round((m * 100)/nrow(GI_standard), 0), "% done"))
}
close(pb)

# WRITE OUTPUT INTEGRATED PLOT BASED GROUND INVENTORY-STANDARD-VOXEL METRICS FILE IN .XLSX FORMAT

GI_standard_voxel_file <- createWorkbook("GI_standard_voxel_metrics")
addWorksheet(GI_standard_voxel_file, "GI_standard_voxel_metrics")
writeData(GI_standard_voxel_file, sheet = "GI_standard_voxel_metrics", GI_standard_voxel)
saveWorkbook(GI_standard_voxel_file,
             "WRITE FOLDER PATH AND OUTPUT INTEGRATED PLOT BASED GROUND INVENTORY-STANDARD-VOXEL METRICS FILE IN .XLSX FORMAT",
             overwrite = TRUE)

end.time = Sys.time()
time.taken = end.time - start.time
print(time.taken)

