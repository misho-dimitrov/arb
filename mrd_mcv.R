# Mihail Dimitrov @ mihail.dimitrov@kcl.ac.uk

# !!! Make sure the raw morphometry tables contain only 8 ADNIs per subjects!
# !!! Make sure that the values in the excel tables are actually formatted as numbers
# !!! Make sure that you perform a quality check on a number of random ICC outputs (by running them again ON THEIR OWN and comparing them to an SPSS output)


# install the libraries, if necessary
# install.packages('readxl')
# install.packages('xlsx')

# load the libraries
library(readxl)
library(xlsx)

# set the dataset
X1 <- read_excel("/Users/mishodimitrov/Downloads/aseg_stats_cross.xlsx", sheet = "Sheet1")

# set a range to include each brain region in the dataset
brain_region <- seq(2, 45, by = 1)


##################################################################################################


# set individual ADNI lists
for (i in brain_region) {
    a1 <- X1[c(1,9,17,25,33,41,49,57,65,73,81,89,97,105,113,121,129,137,145,153),c(i)]
    a2 <- X1[c(2,10,18,26,34,42,50,58,66,74,82,90,98,106,114,122,130,138,146,154),c(i)]
    a3 <- X1[c(3,11,19,27,35,43,51,59,67,75,83,91,99,107,115,123,131,139,147,155),c(i)]
    a4 <- X1[c(4,12,20,28,36,44,52,60,68,76,84,92,100,108,116,124,132,140,148,156),c(i)]
    b1 <- X1[c(5,13,21,29,37,45,53,61,69,77,85,93,101,109,117,125,133,141,149,157),c(i)]

    
    # set the subsets for comparison
    a1a2 = data.frame(v1=a1, v2=a2)
    a1a3 = data.frame(v1=a1, v2=a3)
    a1a4 = data.frame(v1=a1, v2=a4)
    a1b1 = data.frame(v1=a1, v2=b1)
    a3a4 = data.frame(v1=a3, v2=a4)
    # make a list of the subsets to be used as separate objects of the same list
    comp <- list(a1a3, a1a4, a1b1, a3a4)

    
    # create an empty matrix to be filled with results
    matr <- matrix(nrow = 0, ncol = 1)
    
    ################################################################################################
    
    # calculate the mean relative (percentage) difference
    # Only A1vsA2
    rd <- abs((a1a2[, c(1)] - a1a2[, c(2)])/a1a2[, c(1)])
    mrd <- mean(rd)
    mrd_percentage <- mrd*100
    mrd_n <- length(rd)
    mrd_sd <- sd(rd)
    mrd_se <- mrd_sd/sqrt(mrd_n)
    alpha <- 0.05
    mrd_degrees_freedom <- mrd_n - 1
    mrd_t_score <- qt(p=alpha/2, df=mrd_degrees_freedom,lower.tail=F)
    mrd_margin_error <- mrd_t_score * mrd_se
    mrd_lbound <- mrd - mrd_margin_error
    mrd_ubound <- mrd + mrd_margin_error
    mrd_percentage_lbound <- mrd_lbound*100
    mrd_percentage_ubound <- mrd_ubound*100
    matr2 <- rbind(matr, mrd_percentage, mrd_percentage_lbound, mrd_percentage_ubound)
    
    # Calculate the MRD for the rest of the session comparisons
    for (objc in comp) {
        rd <- abs((objc[, c(1)] - objc[, c(2)])/objc[, c(1)])
        mrd <- mean(rd)
        mrd_percentage <- mrd*100
        mrd_n <- length(rd)
        mrd_sd <- sd(rd)
        mrd_se <- mrd_sd/sqrt(mrd_n)
        alpha <- 0.05
        mrd_degrees_freedom <- mrd_n - 1
        mrd_t_score <- qt(p=alpha/2, df=mrd_degrees_freedom,lower.tail=F)
        mrd_margin_error <- mrd_t_score * mrd_se
        mrd_lbound <- mrd - mrd_margin_error
        mrd_ubound <- mrd + mrd_margin_error
        mrd_percentage_lbound <- mrd_lbound*100
        mrd_percentage_ubound <- mrd_ubound*100
      
        mrd_all <- c(mrd_percentage, mrd_percentage_lbound, mrd_percentage_ubound)
        matr2 <- cbind(matr2, mrd_all)
    }
    
    # name the columns and export the entire dataset to an excel file
    colnames(matr2) <- c("A1 vs A2", "A1 vs A3", "A1 vs A4", "A1 vs B1", "A3 vs A4")
    main_path <- path.expand("/Users/mishodimitrov/Downloads/")
    region_name <- colnames(a1)
    write.xlsx(matr2, paste0(main_path, "MRD.xlsx"), sheetName = region_name, col.names = TRUE, row.names = TRUE, append = TRUE)
    
    
    ##################################################################################################
    
    # create a new empty matrix for the MCV results
    matr_a <- matrix(nrow = 0, ncol = 1)
    
    # calculate the mean coefficient of variation
    # Only A1vsA2
    df_1d <- as.vector(t(a1a2))
    mcv_mean <- mean(df_1d)
    mcv_sd <- sd(df_1d)
    mcv_value <- mcv_sd/mcv_mean
    mcv_percentage <- mcv_value*100
    mcv_percentage_lbound <- mcv_percentage
    mcv_percentage_ubound <- mcv_percentage
    matr_b <- rbind(matr_a, mcv_percentage, mcv_percentage_lbound, mcv_percentage_ubound)
    
    # Calculate the MCV for the rest of the session comparisons
    for (objc in comp) {
      df_1d <- as.vector(t(objc))
      mcv_mean <- mean(df_1d)
      mcv_sd <- sd(df_1d)
      mcv_value = mcv_sd/mcv_mean
      mcv_percentage = mcv_value*100
      mcv_percentage_lbound <- mcv_percentage
      mcv_percentage_ubound <- mcv_percentage
      mcv_all <- c(mcv_percentage, mcv_percentage_lbound, mcv_percentage_ubound)
      matr_b <- cbind(matr_b, mcv_all)
    }
    
    
    # name the columns and export the entire dataset to an excel file
    colnames(matr_b) <- c("A1 vs A2", "A1 vs A3", "A1 vs A4", "A1 vs B1", "A3 vs A4")
    main_path <- path.expand("/Users/mishodimitrov/Downloads/")
    region_name <- colnames(a1)
    write.xlsx(matr_b, paste0(main_path, "MCV.xlsx"), sheetName = region_name, col.names = TRUE, row.names = TRUE, append = TRUE)
    
}
    
print("Done.")
