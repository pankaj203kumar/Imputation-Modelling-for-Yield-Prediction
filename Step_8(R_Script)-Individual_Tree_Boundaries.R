# THIS CODE DETECTS INDIVIDUAL TREE BOUNDARIES & LOCATION (XYZ) INFORMATION FROM LIDAR DERVIED CANOPY HEIGHT MODEL (CHM) USING rLiDAR PACKAGE
# -----------------------------------------------

rm(list = ls()) #REMOVES ALL THE VARIABLES
cat("\f") #CLEAR SCREEN

setwd("SET WORKING DIRECTORY")

require(rLiDAR)
require(lidR)
require(raster)
require(rgeos)
require(rgdal)
require(sp)
require(tcltk)

input_folder <- "SET INPUT FOLDER WHERE THE CHM FILES IN .img FORMAT ARE LOCATED"
files_list <- list.files(path = input_folder, pattern = "*.img") #CHM FILES IN .img FORMAT
no_files <- length(files_list)

progress_bar <- tkProgressBar(title = "Progress Bar", min = 1, max = no_files, width = 300) #SET UP THE PROGRESS BAR
b <- 0
for (a in files_list)
{
  input_file <- a
  file_name <- strsplit(input_file, split = ".img")
  
  # APPLY SMOOTHING TO LIDAR DERIVED CHM TO REMOVE SPURIOUS LOCAL MAXIMA POINTS
  filter <- "mean" # MEAN FILTER
  ws <- 7 # WINDOW SIZE
  chm_file <- raster(input_file, band = 1)
  chm_smooth <- CHMsmoothing(chm_file, filter, ws) # APPLIES SMOOTHING TO CHM USING MEAN FILTER AND WINDOW SIZE OF 7
  
  # DETECT THE POSITION (XYZ) OF EACH SINGLE TREE WITHIN THE CHM
  
  fws <- 7 #SQUARE WINDOW SIZE
  minht <- 1.37 #HEIGHT THRESHOLD
  single_tree_position <- FindTreesCHM(chm_smooth, fws, minht) # FINDS A SINGLE TREE LOCATION BASED ON LOCAL MAXIMA...
  # ...VARIATION WITHIN A FIXED WINDOW SIZE & RETURNS 4 COLUMN MATRIX (TREE_ID,X,Y,Z)
  
  # DETECT THE BOUNDARY OF EACH SINGLE TREE FROM CHM
  
  max_crown <- 5.0 #MAXIMUM INDIVIDUAL TREE CROWN RADIUS
  exclusion <- 0.2 #SINGLE VALUE VARYING FROM 0-1 THAT REPRESENTS THE PIXEL EXCLUSION. FOR EXAMPLE-THE EXCLUSION...
  # ...VALUE OF 0.2 WILL EXCLUDE ALL THE PIXELS FOR A SINGLE TREE THAT HAS A HEIGHT VALUE OF LESS THAN 20% OF THE...
  # ...MAXIMUM HEIGHT FROM THE SAME TREE.
  
  single_tree_canopy <- ForestCAS(chm_smooth, single_tree_position, max_crown, exclusion) #DELINEATES SINGLE TREE...
  # ...BOUNDARY BASED ON MAXIMUM CROWN RADIUS AND HEIGHT VARIATION. IT RETURNS SINGLE TREE BOUNDARY POLYGON AND...
  # ...4-COLUMN MATRIX (X,Y,Z, GROUND PROJECTED BOUNDARY AREA IN SQUARE METRES
  
  # WRITE BOUNDARY OF EACH SINGLE TREE IN ESRI .SHP FORMAT
  
  single_tree_canopy_boundary <- single_tree_canopy[[1]]
  single_tree_canopy_boundary_file_name <- paste(file_name, "-Canopy_Boundary")
  
  writeOGR(single_tree_canopy_boundary, dsn = 'Ind_Canopy_Sel_ShapeF-rLiDAR', layer = single_tree_canopy_boundary_file_name,
           driver = 'ESRI Shapefile')
  
  # WRITE SINGLE TREE LOCATION, HEIGHT & CANOPY AREA IN .CSV FORMAT
  
  single_tree_canopy_position <- single_tree_canopy[[2]]
  single_tree_canopy_position_file_name <- paste(file_name, "-Canopy_Position")
  
  csv_file_name <- paste(single_tree_canopy_position_file_name, "csv", sep=".") # IN .CSV FORMAT
  
  write.csv(single_tree_canopy_position, csv_file_name)
  
  Sys.sleep(0.1)
  b <- b + 1
  setTkProgressBar(progress_bar, a, label = paste(round(b/no_files) * 100, 0), "% done")
}
close(progress_bar)
