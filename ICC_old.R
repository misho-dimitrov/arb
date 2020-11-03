# Mihail Dimitrov @ mihail.dimitrov@kcl.ac.uk

# !!! Make sure the raw morphometry tables contain only 8 ADNIs per subjects!
# !!! Make sure that the values in the excel tables are actually formatted as numbers
# !!! Make sure that you perform a quality check on a number of random ICC outputs (by running them again ON THEIR OWN)


# install the libraries, if necessary
# install.packages('readxl')
# install.packages('xlsx')
# install.packages('irr')

# load the libraries
library(readxl)
library(xlsx)
library(irr)

# set the dataset
X1 <- read_excel("~/Downloads/1.xlsx", sheet = "Sheet1")

# set a range to include each brain region in the dataset
brain_region <- seq(2, 64, by = 1)

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

    # create an empty matrix to be filled with ICC results
    matr <- matrix(nrow = 0, ncol = 1)

    #set the rows
    icc_test <- (icc(a1a2, model = "twoway", type = "agreement", unit = "single", r0 = 0, conf.level = 0.95))
    icc_value <- icc_test$value
    icc_lbound <- icc_test$lbound
    icc_ubound <- icc_test$ubound
    icc_p_value <- icc_test$p.value
    matr2 <- rbind(matr, icc_value, icc_lbound, icc_ubound, icc_p_value)


    # go through each comparison (2 onwards), perform ICC and append columns to matrix
    for (objc in comp) {
        icc_test <- (icc(objc, model = "twoway", type = "agreement", unit = "single", r0 = 0, conf.level = 0.95))
        icc_value <- icc_test$value
        icc_lbound <- icc_test$lbound
        icc_ubound <- icc_test$ubound
        icc_p_value <- icc_test$p.value
        icc_all <- c(icc_value, icc_lbound, icc_ubound, icc_p_value)
        matr2 <- cbind(matr2, icc_all)
        
    }
    
    # name the columns and export the entire ICC dataset to an excel file
    colnames(matr2) <- c("A1 vs A2", "A1 vs A3", "A1 vs A4", "A1 vs B1", "A3 vs A4")
    main_path <- path.expand("~/Downloads/")
    region_name <- colnames(a1)
    # write.csv(matr2, paste0(main_path, "ICC.xlsx", na = "NA", append = TRUE, col_names = TRUE))
    write.xlsx(matr2, paste0(main_path, "ICC.xlsx"), sheetName = region_name, col.names = TRUE, row.names = TRUE, append = TRUE)
    
}
    
# to do:    
# plot each spreadsheet as a single entity in a graph (ICC + CI + p-value)
# save plot as an image file?