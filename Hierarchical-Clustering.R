# ========== Load packages ==========
rm(list = ls())
library(readxl)
library(writexl)
library(grid)
library(ComplexHeatmap)
library(circlize)
library(viridis)

# Data directory (input and output)
base_dir <- "D:/Feature_Importance_Data/"

# Define seasons and their clustering parameters
seasons <- list(
  Annual = list(k = 5, dist_method = "euclidean", hclust_method = "average"),
  Spring = list(k = 6, dist_method = "canberra",   hclust_method = "complete"),
  Summer = list(k = 3, dist_method = "maximum",    hclust_method = "ward.D2"),
  Autumn = list(k = 6, dist_method = "manhattan",  hclust_method = "average"),
  Winter = list(k = 7, dist_method = "manhattan",  hclust_method = "complete")
)

# Loop over each season
for (season_name in names(seasons)) {
  
  # Input file path
  file_path <- file.path(base_dir, paste0(season_name, ".xlsx"))
  
  # Skip if file does not exist
  if (!file.exists(file_path)) {
    warning(paste("File not found, skipped:", file_path))
    next
  }
  
  # Read and preprocess data
  data <- read_excel(file_path)
  mat <- as.matrix(data[, -1])
  rownames(mat) <- data[[1]]
  feat_names <- colnames(mat)
  
  # Season name for output prefix
  input_basename <- season_name
  
  # Extract parameters for current season
  k <- seasons[[season_name]]$k
  dist_method <- seasons[[season_name]]$dist_method
  hclust_method <- seasons[[season_name]]$hclust_method
  
  # Hierarchical clustering
  hc <- hclust(dist(mat, method = dist_method), method = hclust_method)
  clusters <- cutree(hc, k)
  
  # ---------- Settings ----------
  # Color function for bubbles
  col_fun <- colorRamp2(seq(min(mat), max(mat), length.out = 100),
                        colorRampPalette(c("#4575B4","#91BFDB","#FEE090","#FC8D59","#D73027"))(100))
  
  # Bubble size matrix (normalized to 0.5~5mm)
  bubble_size <- function(x) 0.5 + (x - min(mat)) / diff(range(mat)) * 4.5
  sz <- bubble_size(mat)
  
  # Cluster colors
  if (k <= 8) {
    cluster_cols <- setNames(
      c("#fb9a99","#b2df8a","#92d6dd","#cab2d6","#eebc88","#8CC0DC","#F7CDE0","#A5CAD1")[1:k],
      1:k
    )
  } else {
    cluster_cols <- setNames(viridis(k), 1:k)
  }
  
  # ---------- Build heatmap ----------
  hm <- Heatmap(mat,
                col = colorRamp2(range(mat), c("white","white")),
                cluster_rows = hc,
                cluster_columns = FALSE,
                row_dend_width = unit(30, "mm"),
                show_row_names = TRUE,
                row_names_gp = gpar(fontsize = 6),
                row_names_side = "left",
                show_column_names = TRUE,
                column_names_gp = gpar(fontsize = 6, rot = 45),
                column_names_side = "bottom",
                width = unit(ncol(mat)*5, "mm"),
                height = unit(nrow(mat)*5, "mm"),
                cell_fun = function(j,i,x,y,w,h,f) {
                  r <- min(convertWidth(w,"mm",T), convertHeight(h,"mm",T)) * 0.4
                  r_act <- min(sz[i,j]/2, r)
                  grid.circle(x, y, r = unit(r_act,"mm"),
                              gp = gpar(fill = col_fun(mat[i,j]), col = col_fun(mat[i,j]), lwd = 0.1))
                },
                show_heatmap_legend = FALSE,
                left_annotation = rowAnnotation(
                  Cluster = factor(clusters, levels = 1:k),
                  col = list(Cluster = cluster_cols),
                  show_legend = FALSE,
                  show_annotation_name = TRUE,
                  width = unit(5,"mm")
                ),
                border = FALSE,
                rect_gp = gpar(col = "#E0E0E0", lwd = 0.5)
  )
  
  # ---------- Legends ----------
  cluster_legend <- Legend(
    labels = paste0("Cluster", 1:k),
    title = "Cluster",
    legend_gp = gpar(fill = cluster_cols),
    labels_gp = gpar(fontsize = 6),
    title_gp = gpar(fontsize = 7),
    row_gap = unit(1, "mm") 
  )
  
  at_values <- round(seq(min(mat), max(mat), length.out = 5), 2) 
  lgd <- Legend(col_fun = col_fun, 
                title = "Relative importance", 
                at = at_values,
                labels = sprintf("%.2f", at_values), 
                title_gp = gpar(fontsize = 7),
                labels_gp = gpar(fontsize = 6), 
                legend_height = unit(25, "mm"))
  
  # Save PDF
  pdf(file.path(base_dir, paste0(input_basename, "_Hierarchical_Clustering_Heatmap.pdf")), 6, 8)
  draw(hm, 
       annotation_legend_list = list(cluster_legend, lgd), 
       heatmap_legend_side = "right",
       annotation_legend_side = "right",
       merge_legend = TRUE, 
       padding = unit(c(10,10,10,10),"mm"))
  dev.off()
  
  # Save PNG
  png(file.path(base_dir, paste0(input_basename, "_Hierarchical_Clustering_Heatmap.png")), 1000, 1200, res = 150)
  draw(hm, 
       annotation_legend_list = list(cluster_legend, lgd),
       heatmap_legend_side = "right",
       annotation_legend_side = "right",
       merge_legend = TRUE,
       padding = unit(c(10,10,10,10),"mm"))
  dev.off()
  
  # ---------- Generate result tables ----------
  clustered <- data.frame(City_ID = rownames(mat), Cluster = paste0("C", clusters), mat)
  means <- aggregate(mat, by = list(Cluster = clusters), FUN = mean)
  means$Cluster <- paste0("C", means$Cluster)
  
  counts <- table(clusters)
  summary_tab <- data.frame(Cluster = paste0("C", 1:k), 
                            n_Cities = as.numeric(counts),
                            Percent = round(100 * as.numeric(counts) / nrow(mat), 1))
  for (i in 1:k) {
    vals <- as.numeric(means[i, -1])
    top_idx <- order(vals, decreasing = TRUE)[1:min(12, length(vals))]
    for (j in seq_along(top_idx)) {
      summary_tab[i, paste0("Top", j, "_Feature")] <- 
        paste0(feat_names[top_idx[j]], " (", round(vals[top_idx[j]], 3), ")")
    }
  }
  
  # Save Excel file
  write_xlsx(list(Clustering_Results = clustered, 
                  Cluster_Means = means, 
                  Cluster_Summary = summary_tab),
             file.path(base_dir, paste0(input_basename, "_Cluster_Analysis_Results.xlsx")))
  
  cat("Completed:", season_name, "\n")
}
