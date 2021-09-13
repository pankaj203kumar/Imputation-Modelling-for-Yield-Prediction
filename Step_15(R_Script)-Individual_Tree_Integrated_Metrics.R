# THIS CODE INTEGRATES INDIVIDUAL TREE BASED GROUND INVENTORY METRICS WITH THE LIDAR DERIVED STANDARD AND VOXEL METRICS BASED ON TREE_IDs FOR EACH PLOT
# -----------------------------------------------

rm(list = ls()) #REMOVES ALL THE VARIABLES
cat("\f") #CLEAR SCREEN

setwd("SET WORKING DIRECTORY")

require(openxlsx)
require(matlib)
require(Metrics)
require(dplyr)

# FIRST, INTEGRATES INDIVIDUAL TREE BASED GROUND INVENTORY METRICS WITH LIDAR DERIVED STANDARD METRICS

GI_file = read.xlsx("SET INPUT GROUND INVENTORY METRICS FILE IN .XLSX FORMAT") #PLOT_ID,TREE_ID AND GROUND INVENTORY METRICS
standard_file = read.xlsx("SET INPUT LIDAR DERIVED STANDARD METRICS FILE IN .XLSX FORMAT") #PLOT_ID,TREE_ID AND LIDAR DERIVED STANDARD METRICS

View(GI_file)
View(standard_file)

GI = data.frame(GI_file)
standard = data.frame(standard_file)

plot_ids = cbind(unique(standard[, 1]))
nrow_plot_ids = nrow(plot_ids)

GI_standard = data.frame()

pb <- winProgressBar(title = "Progress Bar", min = 1, max = nrow_plot_ids, initial = 1, width = 300)
for (m in 1:nrow_plot_ids)
{
  for (n in 1:nrow(standard))
  {
    for (p in 1:nrow(GI))
    {
      if ((standard[n, 1] == plot_ids[m]) & (GI[p, 4] == plot_ids[m]) & (standard[n, 2] == GI[p, 5])) # MATCHES PLOT_IDs AND TREE_IDs FOR BOTH...
        # ...GROUND INVENTORY AND STANDARD FILES
      {
        GI_standard_col = cbind(GI[p, 4:5], GI[p, 6:14], standard[n, 3:68])
        GI_standard = rbind(GI_standard, GI_standard_col)
      }
    }
  }
  Sys.sleep(0.1)
  setWinProgressBar(pb, m, title = paste(round((m * 100)/nrow(nrow_plot_ids), 0), "% done"))
}
close(pb)

GI_standard_file <- createWorkbook("GI_standard_metrics")
addWorksheet(GI_standard_file, "GI_standard_metrics")
writeData(GI_standard_file, sheet = "GI_standard_metrics", GI_standard)
saveWorkbook(GI_standard_file,
             "WRITE FOLDER PATH AND OUTPUT INTEGRATED INDIVIDUAL TREE BASED GROUND INVENTORY-STANDARD METRICS FILE IN .XLSX FORMAT",
             overwrite = TRUE)

# SECOND, INTEGRATES INDIVIDUAL TREE BASED GROUND INVENTORY-STANDARD METRICS WITH LIDAR DERIVED VOXEL METRICS

voxel_file = read.xlsx("SET INPUT VOXEL METRICS FILE IN .XLSX FORMAT") #PLOT_ID,TREE_ID AND LIDAR DERIVED VOXEL METRICS

View(voxel_file)
View(GI_standard_file)

voxel = data.frame(voxel_file)
GI_standard = data.frame(GI_standard_file)

GI_standard_voxel = data.frame()
SSSSSS
pb <- winProgressBar(title = "Progress Bar", min = 1, max = nrow_plot_ids, initial = 1, width = 300)
for (m in 1:nrow_plot_ids)
{
  for (n in 1:nrow(GI_standard))
  {
    for (p in 1:nrow(voxel))
    {
      if ((GI_standard[n, 1] == plot_ids[m]) & (voxel[p, 1] == plot_ids[m]) & (GI_standard[n, 2] == voxel[p, 2])) # MATCHES PLOT_IDs AND TREE_IDs FOR BOTH...
        # ...GROUND INVENTORY AND VOXEL FILES
      {
        GI_standard_voxel_col = cbind(GI_standard[n, 1:2], GI_standard[n, 3:77], voxel[p, 3:153])
        GI_standard_voxel = rbind(GI_standard_voxel, GI_standard_voxel_col)
      }
    }
  }
  Sys.sleep(0.1)
  setWinProgressBar(pb, m, title = paste(round((m * 100)/nrow(nrow_plot_ids), 0), "% done"))
}
close(pb)

# WRITE OUTPUT INTEGRATED INDIVIDUAL TREE BASED GROUND INVENTORY-STANDARD-VOXEL METRICS FILE IN .XLSX FORMAT

GI_standard_voxel_file <- createWorkbook("GI_standard_voxel_metrics")
addWorksheet(GI_standard_voxel_file, "GI_standard_voxel_metrics")
writeData(GI_standard_voxel_file, sheet = "GI_standard_voxel_metrics", GI_standard_voxel)
saveWorkbook(GI_standard_voxel_file,
             "WRITE FOLDER PATH AND OUTPUT INTEGRATED INDIVIDUAL TREE BASED GROUND INVENTORY-STANDARD-VOXEL METRICS FILE IN .XLSX FORMAT",
             overwrite = TRUE)
