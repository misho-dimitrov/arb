# Mihail Dimitrov @ mihail.dimitrov@kcl.ac.uk

# !!! Make sure the raw morphometry tables contain only 8 ADNIs per subjects!
# !!! Make sure that the values in the excel tables are actually formatted as numbers
# !!! Make sure that you perform a quality check on a number of random ICC outputs (by running them again ON THEIR OWN)


# install the libraries, if necessary
# install.packages('readxl')
# install.packages('xlsx')
# install.packages('irr')

install.packages('readxl')
install.packages('xlsx')
install.packages('irr')

# load the libraries
library(readxl)
library(xlsx)
library(irr)

# set the dataset
## X1 <- read_excel("~/Downloads/1.xlsx", sheet = "sheet1")
X1 <- read_excel("D:/PhD Psychosis Study/PIN study/FS_ICC_BRCPIN_06042021/CorticalVolume_BRCPIN_FS7_long_rh.xlsx", sheet = "Sheet1")

# set a range to include each brain region in the dataset
brain_region <- seq(2, 35, by = 1)
    
# set individual ADNI lists 
for (i in brain_region) {
    a1 <- X1[c(1,5,9,13,17,21,25,29,33,37,41,45,49,53,57,61,65,69,73,77),c(i)]
    a2 <- X1[c(2,6,10,14,18,22,26,30,34,38,42,46,50,54,58,62,66,70,74,78),c(i)]
    b1 <- X1[c(3,7,11,15,19,23,27,31,35,39,43,47,51,55,59,63,67,71,75,79),c(i)]       
    b2 <- X1[c(4,8,12,16,20,24,28,32,36,40,44,48,52,56,60,64,68,72,86,80),c(i)]
        

    # set the subsets for comparison
    a1a2 = data.frame(V1=a1, V2=a2)
    b1b2 = data.frame(V1=b1, V2=b2)
    a1b1 = data.frame(V1=a1, V2=b1)
    a2b2 = data.frame(V1=a2, V2=b2)
    # make a list of the subsets to be used as separate objects of the same list
    comp <- list(b1b2, a1b1, a2b2)

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
    colnames(matr2) <- c("C1 vs C2", "D1 vs D2", "C1 vs D1", "C2 vs D2")
    main_path <- path.expand("D:/PhD Psychosis Study/PIN study/FS_ICC_BRCPIN_06042021")
    region_name <- colnames(a1)
    # write.csv(matr2, paste0(main_path, "hippocampal_subfields_pin_cross_ICC.xlsx", na = "NA", append = TRUE, col_names = TRUE))
    write.xlsx(matr2, paste0("CorticalVolume_BRCPIN_FS7_long_rh_ICC.xlsx"), sheetName = region_name, col.names = TRUE, row.names = TRUE, append = TRUE)
    
}
    
# to do:    
# plot each spreadsheet as a single entity in a graph (ICC + CI + p-value)
# save plot as an image file?