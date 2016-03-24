# load text mining library
library(tm)
# this library is good for opening up large xlsx files
library(openxlsx)

filename <- "LAN_SWITCHING_SRs.xlsx"

SR <- read.xlsx(filename, sheet = 1, colNames = TRUE)
# Beware: over 272,000 rows

# Only look at 20000 rows
SR_short <- SR[sample(seq(1, nrow(SR), by = 1), 20000),]

SR_text <- paste(SR_short$SR.Problem.Summary, collapse = " ")

SR_source <- VectorSource(SR_text)
corpus <- Corpus(SR_source)

# Clean up text
corpus <- tm_map(corpus, content_transformer(tolower))
corpus <- tm_map(corpus, removePunctuation)
corpus <- tm_map(corpus, stripWhitespace)
corpus <- tm_map(corpus, removeWords, stopwords(kind = "en"))

# Create document-term matrix of 1 column
dtm <- DocumentTermMatrix(corpus)

# Terms that appear at least 100 times (sorted alphabetically)
findFreqTerms(dtm, 100)

### Convert dtm to regular matrix and find frequency of all terms
dtm2 <- as.matrix(dtm)
frequency <- colSums(dtm2)
frequency <- sort(frequency, decreasing = TRUE)

### Top 100 words
frequency[1:100]

#Resources:
#https://deltadna.com/blog/text-mining-in-r-for-term-frequency/
#https://cran.r-project.org/web/packages/tm/vignettes/tm.pdf