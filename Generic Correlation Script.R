# Prerequisites:
# Excel file with the following sheets:
#     Sheet1: Component names in first row and data below (component data organized in columns)
#     Sheet2: 2 columns: Component and Product (first row has these names)

## Instructions to generate Sheet1:
# 1.) Export component data.
##### Do this for one product at a time to generate Sheet2 easily.
# 2.) Filter exported file to only show number of bugs for each component
# 3.) Copy data for all components and transpose-paste into a new sheet named "Sheet1"
# 4.) Delete column A and rows 2-4. Only do this step after all products are entered.

## Instructions to generate Sheet2:
# 1.) From the filtered sheet of Step 2 of the Sheet1 instructions,
#     copy the components to the first column
#     of a new sheet named "Sheet2". Label this column "Component".
# 2.) In column B, create a column called "Product" and fill it with the product name.
# WARNING: There may be duplicate component names that you need to resolve.

## Instructions for adding additional products:
# 1.) Repeat the Sheet1 and Sheet2 instructions for each product,
#     appending the new information to the old.


library(xlsx)
library(corrplot)  # allows correlation visuals

# IMPORTANT:
#Specify name of Excel file here and make sure it is in the working directory.
#Specify minimum number of SRs needed for inclusion in analysis
filename <- "ComponentData.xlsx"
threshold <- 96

#read in data
rawdat <- read.xlsx(filename, sheetName = "Sheet1")
dat <- rawdat[names(which(colSums(rawdat) > threshold))]

#a mapping of components to their product
prodCompMap <- read.xlsx(filename, sheetName = "Sheet2")

#replace missing values with 0
dat[is.na(dat)] <- 0

#generate correlation matrix
felix <- cor(dat)

#if small analysis, plot correlation matrix
if (ncol(dat) < 25) {
  corrplot(felix, method = "circle", type = "lower", diag = FALSE,
           title = "Correlations between Software Components")
}

felix[lower.tri(felix, diag=TRUE)] = NA  #Prepare to drop duplicates and 1s
rawCorrRank = as.data.frame(as.table(felix))  #Turn into a 3-column table
corrRank = na.omit(rawCorrRank)  #Get rid of duplicates and 1s
colnames(corrRank) <- c("Comp1", "Comp2", "Correlation")

#Replace periods with hyphens for matching later
corrRank$Comp1 <- gsub(".", "-", corrRank$Comp1, fixed = TRUE)
corrRank$Comp2 <- gsub(".", "-", corrRank$Comp2, fixed = TRUE)

corrRank <- merge(corrRank, prodCompMap, by.x = "Comp1", by.y = "Component", all.x = TRUE)
corrRank <- merge(corrRank, prodCompMap, by.x = "Comp2", by.y = "Component", all.x = TRUE)
corrRank <- corrRank[c(4,2,5,1,3)]
colnames(corrRank) <- c("Product1", "Component1", "Product2", "Component2", "Correlation")

corrRankSorted <- corrRank[order(corrRank$Correlation, decreasing = TRUE),]    #Sort by highest positive correlation
diffProd <- corrRankSorted[corrRankSorted$Product1!=corrRankSorted$Product2,]

# Show top 6 correlations
head(diffProd)

# Other exploration
tail(diffProd)
summary(diffProd$Correlation)
