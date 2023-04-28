import os
import re
import lexnlp.extract.en.money
import pandas as pd

from PyPDF2 import PdfReader
from nltk import tokenize

# Constants
ENVIROMENTAL_KEYWORDS = ["Material", "Energy", "Water", "Biodiversity", "Emission", "Effluent", "Waste", "Climate",
                         "Paper", "Compliance", "Transport", "Environment", "Environmental initiatives",
                         "Environmental assessment", "Environmental operations"]

SOCIAL_KEYWORDS = ["Employment", "Labor management", "Employee", "Health", "Safety", "Training", "Education",
                   "Diversity", "Woman management", "Equal opportunity", "Equal remuneration", "Employee grievance",
                   "Discrimination", "Freedom of association", "Collective bargaining", "Child labor", "Forced labor", "Compulsory labor",
                   "Security", "Indigenous rights", "Human rights", "Public policy", "Corruption", "Anti-competitive behaviour",
                   "Compliance", "Society", "Impact society", "Local communities", "Labelling", "Marketing Communication",
                   "Customer privacy", "Customer compliance", "Economic impact", "Procurement practices", "Pay Ratio",
                   "Employee turnover", "Temporary worker", "Injury rate", "Labour"]

GOVERNANCE_KEYWORDS = ["Board Diversity", "Board Independence", "Incentives", "Collective Bargaining", "Code of Conduct",
                       "Ethics", "Corruption", "Data Privacy", "ESG Reporting", "Disclosure practices", "External Assurance",
                       "Data Protection", "Fair remuneration", "Independent Director", "Board of Directors", "Board Meeting",
                       "Committee"]

PDF_PAGES_TO_INCLUDE = [
    {"file_name": "4689 JP.pdf", "page_start": 56, "page_end": 88},
    {"file_name": "6506 JP.pdf", "page_start": 26, "page_end": 41},
    {"file_name": "WLN FP.pdf", "page_start": 40, "page_end": 86},
    {"file_name": "9021 JP.pdf", "page_start": 35, "page_end": 49},
    {"file_name": "3141 JP.pdf", "page_start": 16, "page_end": 22},
    {"file_name": "VOLCARB SS.pdf", "page_start": 147, "page_end": 199},
    {"file_name": "VOLVB SS.pdf", "page_start": 147, "page_end": 199},
    {"file_name": "VOLVA SS.pdf", "page_start": 147, "page_end": 199},
    {"file_name": "VER AV.pdf", "page_start": 31, "page_end": 49},
    {"file_name": "VER AV.pdf", "page_start": 174, "page_end": 213},
    {"file_name": "4732 JP.pdf", "page_start": 17, "page_end": 26},
    {"file_name": "UPM FH.pdf", "page_start": 30, "page_end": 59},
    {"file_name": "UCG IM.pdf", "page_start": 74, "page_end": 119},
    {"file_name": "UCB BB.pdf", "page_start": 25, "page_end": 121},
    {"file_name": "8015 JP.pdf", "page_start": 40, "page_end": 77},
    {"file_name": "5332 JP.pdf", "page_start": 14, "page_end": 42},
    {"file_name": "9005 JP.pdf", "page_start": 27, "page_end": 41},
    {"file_name": "9501 JP.pdf", "page_start": 26, "page_end": 63},
    {"file_name": "9501 JP.pdf", "page_start": 88, "page_end": 93},
    {"file_name": "9602 JP.pdf", "page_start": 16, "page_end": 36},
    {"file_name": "3626 JP.pdf", "page_start": 10, "page_end": 46},
    {"file_name": "HO FP.pdf", "page_start": 16, "page_end": 47},
    {"file_name": "TRN IM.pdf", "page_start": 167, "page_end": 274},
    {"file_name": "TELIA SS.pdf", "page_start": 54, "page_end": 127},
    {"file_name": "TEL2B SS.pdf", "page_start": 39, "page_end": 82},
    {"file_name": "4502 JP.pdf", "page_start": 28, "page_end": 55},
    {"file_name": "4502 JP.pdf", "page_start": 61, "page_end": 65},
    {"file_name": "SWEDA SS.pdf", "page_start": 201, "page_end": 260},
    {"file_name": "SHBA SS.pdf", "page_start": 36, "page_end": 91},
    {"file_name": "SCAB SS.pdf", "page_start": 140, "page_end": 173},
    {"file_name": "5713 JP.pdf", "page_start": 21, "page_end": 71},
    {"file_name": "SOLB BB.pdf", "page_start": 55, "page_end": 64},
    {"file_name": "GLE FP.pdf", "page_start": 15, "page_end": 25},
    {"file_name": "GLE FP.pdf", "page_start": 32, "page_end": 38},
    {"file_name": "SKAB SS.pdf", "page_start": 40, "page_end": 93},
    {"file_name": "SEBA SS.pdf", "page_start": 32, "page_end": 65},
    {"file_name": "5831 JP.pdf", "page_start": 27, "page_end": 67},
    {"file_name": "4507 JP.pdf", "page_start": 50, "page_end": 89},
    {"file_name": "7701 JP.pdf", "page_start": 30, "page_end": 52},
    {"file_name": "SGSN SW.pdf", "page_start": 36, "page_end": 124},
    {"file_name": "9143 JP.pdf", "page_start": 28, "page_end": 35},
    {"file_name": "SECUB SS.pdf", "page_start": 27, "page_end": 38},
    {"file_name": "SECUB SS.pdf", "page_start": 132, "page_end": 147},
    {"file_name": "8473 JP.pdf", "page_start": 17, "page_end": 49},
    {"file_name": "SAP GR.pdf", "page_start": 271, "page_end": 325},
    {"file_name": "SAF FP.pdf", "page_start": 7, "page_end": 24},
    {"file_name": "6963 JP.pdf", "page_start": 2, "page_end": 27},
    {"file_name": "8308 JP.pdf", "page_start": 8, "page_end": 37},
    {"file_name": "REP SM.pdf", "page_start": 21, "page_end": 24},
    {"file_name": "REP SM.pdf", "page_start": 51, "page_end": 120},
    {"file_name": "RCO FP.pdf", "page_start": 40, "page_end": 61},
    {"file_name": "RF US.pdf", "page_start": 53, "page_end": 91},
    {"file_name": "2181 JP.pdf", "page_start": 6, "page_end": 31},
    {"file_name": "4578 JP.pdf", "page_start": 28, "page_end": 75},
    {"file_name": "4768 JP.pdf", "page_start": 28, "page_end": 61},
    {"file_name": "9532 JP.pdf", "page_start": 10, "page_end": 41},
    {"file_name": "ORA FP.pdf", "page_start": 57, "page_end": 101},
    {"file_name": "6645 JP.pdf", "page_start": 70, "page_end": 113},
    {"file_name": "3861 JP.pdf", "page_start": 2, "page_end": 21},
    {"file_name": "3861 JP.pdf", "page_start": 32, "page_end": 50},
    {"file_name": "9007 JP.pdf", "page_start": 14, "page_end": 25},
    {"file_name": "4684 JP.pdf", "page_start": 2, "page_end": 16},
    {"file_name": "1802 JP.pdf", "page_start": 7, "page_end": 40},
    {"file_name": "NOVN SW.pdf", "page_start": 17, "page_end": 78},
    {"file_name": "6988 JP.pdf", "page_start": 10, "page_end": 26},
    {"file_name": "9843 JP.pdf", "page_start": 16, "page_end": 32},
    {"file_name": "2002 JP.pdf", "page_start": 15, "page_end": 30},
    {"file_name": "4021 JP.pdf", "page_start": 10, "page_end": 42},
    {"file_name": "9101 JP.pdf", "page_start": 1, "page_end": 39},
    {"file_name": "4516 JP.pdf", "page_start": 12, "page_end": 21},
    {"file_name": "4516 JP.pdf", "page_start": 29, "page_end": 39},
    {"file_name": "4091 JP.pdf", "page_start": 30, "page_end": 47},
    {"file_name": "4612 JP.pdf", "page_start": 40, "page_end": 56},
    {"file_name": "6594 JP.pdf", "page_start": 13, "page_end": 34},
    {"file_name": "MOWI NO.pdf", "page_start": 25, "page_end": 145},
    {"file_name": "8593 JP.pdf", "page_start": 10, "page_end": 27},
    {"file_name": "6479 JP.pdf", "page_start": 30, "page_end": 41},
    {"file_name": "2269 JP.pdf", "page_start": 25, "page_end": 31},
    {"file_name": "LR FP.pdf", "page_start": 12, "page_end": 21},
    {"file_name": "FDJ FP.pdf", "page_start": 17, "page_end": 25},
    {"file_name": "6971 JP.pdf", "page_start": 6, "page_end": 27},
    {"file_name": "KPN NA.pdf", "page_start": 17, "page_end": 61},
    {"file_name": "3635 JP.pdf", "page_start": 16, "page_end": 25},
    {"file_name": "4967 JP.pdf", "page_start": 32, "page_end": 48},
    {"file_name": "2503 JP.pdf", "page_start": 20, "page_end": 104},
    {"file_name": "9041 JP.pdf", "page_start": 26, "page_end": 50},
    {"file_name": "9433 JP.pdf", "page_start": 18, "page_end": 53},
    {"file_name": "1812 JP.pdf", "page_start": 26, "page_end": 40},
    {"file_name": "2914 JP.pdf", "page_start": 15, "page_end": 57},
    {"file_name": "8697 JP.pdf", "page_start": 21, "page_end": 35},
    {"file_name": "9201 JP.pdf", "page_start": 64, "page_end": 105},
    {"file_name": "4739 JP.pdf", "page_start": 31, "page_end": 45},
    {"file_name": "2593 JP.pdf", "page_start": 3, "page_end": 33},
    {"file_name": "INW IM.pdf", "page_start": 37, "page_end": 58},
    {"file_name": "5019 JP.pdf", "page_start": 32, "page_end": 61},
    {"file_name": "4062 JP.pdf", "page_start": 19, "page_end": 29},
    {"file_name": "6465 JP.pdf", "page_start": 21, "page_end": 36},
    {"file_name": "HOLN SW.pdf", "page_start": 42, "page_end": 77},
    {"file_name": "6305 JP.pdf", "page_start": 21, "page_end": 35},
    {"file_name": "6965 JP.pdf", "page_start": 14, "page_end": 16},
    {"file_name": "2433 JP.pdf", "page_start": 57, "page_end": 84},
    {"file_name": "3769 JP.pdf", "page_start": 17, "page_end": 21},
    {"file_name": "GFC FP.pdf", "page_start": 14, "page_end": 20},
    {"file_name": "F US.pdf", "page_start": 10, "page_end": 99},
    {"file_name": "ETR US.pdf", "page_start": 23, "page_end": 38},
    {"file_name": "LLY US.pdf", "page_start": 8, "page_end": 13},
    {"file_name": "EDEN FP.pdf", "page_start": 42, "page_end": 53},
    {"file_name": "9020 JP.pdf", "page_start": 18, "page_end": 55},
    {"file_name": "P911 GR.pdf", "page_start": 38, "page_end": 60},
    {"file_name": "4324 JP.pdf", "page_start": 26, "page_end": 55},
    {"file_name": "6902 JP.pdf", "page_start": 17, "page_end": 41},
    {"file_name": "BN FP.pdf", "page_start": 14, "page_end": 32},
    {"file_name": "7912 JP.pdf", "page_start": 28, "page_end": 43},
    {"file_name": "ACA FP.pdf", "page_start": 19, "page_end": 29},
    {"file_name": "CCH LN.pdf", "page_start": 57, "page_end": 81},
    {"file_name": "CCEP US.pdf", "page_start": 26, "page_end": 145},
    {"file_name": "CLX US.pdf", "page_start": 52, "page_end": 70},
    {"file_name": "9502 JP.pdf", "page_start": 53, "page_end": 82},
    {"file_name": "8331 JP.pdf", "page_start": 11, "page_end": 40},
    {"file_name": "9022 JP.pdf", "page_start": 20, "page_end": 39},
    {"file_name": "CLNX SM.pdf", "page_start": 78, "page_end": 229},
    {"file_name": "9697 JP.pdf", "page_start": 53, "page_end": 75},
    {"file_name": "BVI FP.pdf", "page_start": 97, "page_end": 216},
    {"file_name": "BFB US.pdf", "page_start": 20, "page_end": 33},
    {"file_name": "BATS LN.pdf", "page_start": 30, "page_end": 181},
    {"file_name": "5108 JP.pdf", "page_start": 34, "page_end": 50},
    {"file_name": "EN FP.pdf", "page_start": 12, "page_end": 34},
    {"file_name": "BNP FP.pdf", "page_start": 26, "page_end": 37},
    {"file_name": "BCE CN.pdf", "page_start": 24, "page_end": 73},
    {"file_name": "4503 JP.pdf", "page_start": 34, "page_end": 81},
    {"file_name": "G IM.pdf", "page_start": 40, "page_end": 100},
    {"file_name": "ASML NA.pdf", "page_start": 70, "page_end": 151},
    {"file_name": "ASM NA.pdf", "page_start": 56, "page_end": 133},
    {"file_name": "AMS SM.pdf", "page_start": 75, "page_end": 134},
    {"file_name": "7259 JP.pdf", "page_start": 59, "page_end": 142},
    {"file_name": "AC FP.pdf", "page_start": 26, "page_end": 28},
    {"file_name": "XRO AU.pdf", "page_start": 20, "page_end": 30},
    {"file_name": "WKL NA.pdf", "page_start": 36, "page_end": 109},
    {"file_name": "WES AU.pdf", "page_start": 40, "page_end": 92},
    {"file_name": "SOL AU.pdf", "page_start": 36, "page_end": 47},
    {"file_name": "WRT1V FH.pdf", "page_start": 31, "page_end": 119},
    {"file_name": "WDP BB.pdf", "page_start": 14, "page_end": 64},
    {"file_name": "VOD LN.pdf", "page_start": 34, "page_end": 115},
    {"file_name": "VRSN US.pdf", "page_start": 25, "page_end": 35},
    {"file_name": "UMG NA.pdf", "page_start": 126, "page_end": 154},
    {"file_name": "ULVR LN.pdf", "page_start": 27, "page_end": 42},
    {"file_name": "ULVR LN.pdf", "page_start": 60, "page_end": 133},
    {"file_name": "URW NA.pdf", "page_start": 38, "page_end": 235},
    {"file_name": "UMI BB.pdf", "page_start": 73, "page_end": 119},
    {"file_name": "UBI FP.pdf", "page_start": 45, "page_end": 178},
    {"file_name": "4042 JP.pdf", "page_start": 1, "page_end": 37},
    {"file_name": "4042 JP.pdf", "page_start": 48, "page_end": 50},
    {"file_name": "TSCO LN.pdf", "page_start": 18, "page_end": 104},
    {"file_name": "TEMN SW.pdf", "page_start": 32, "page_end": 124},
    {"file_name": "TEL NO.pdf", "page_start": 22, "page_end": 74},
    {"file_name": "TEF SM.pdf", "page_start": 232, "page_end": 760},
    {"file_name": "1972 HK.pdf", "page_start": 88, "page_end": 125},
    {"file_name": "SOBI SS.pdf", "page_start": 92, "page_end": 128},
    {"file_name": "SUN AU.pdf", "page_start": 14, "page_end": 51},
    {"file_name": "STERV FH.pdf", "page_start": 19, "page_end": 119},
    {"file_name": "SGP AU.pdf", "page_start": 35, "page_end": 106},
    {"file_name": "SOON SW.pdf", "page_start": 203, "page_end": 287},
    {"file_name": "SHL AU.pdf", "page_start": 53, "page_end": 66},
    {"file_name": "SOF BB.pdf", "page_start": 27, "page_end": 49},
    {"file_name": "SNA US.pdf", "page_start": 15, "page_end": 35},
    {"file_name": "SKFB SS.pdf", "page_start": 92, "page_end": 150},
    {"file_name": "SPG US.pdf", "page_start": 10, "page_end": 16},
    {"file_name": "SIGN SW.pdf", "page_start": 2, "page_end": 30},
    {"file_name": "SIGN SW.pdf", "page_start": 178, "page_end": 238},
    {"file_name": "SIGN SW.pdf", "page_start": 361, "page_end": 410},
    {"file_name": "ST US.pdf", "page_start": 2, "page_end": 6},
    {"file_name": "SGRO LN.pdf", "page_start": 77, "page_end": 150},
    {"file_name": "SK FP.pdf", "page_start": 49, "page_end": 57},
    {"file_name": "SDR LN.pdf", "page_start": 30, "page_end": 104},
    {"file_name": "DIM FP.pdf", "page_start": 40, "page_end": 124},
    {"file_name": "SRT3 GR.pdf", "page_start": 93, "page_end": 164},
    {"file_name": "SAP CN.pdf", "page_start": 6, "page_end": 12},
    {"file_name": "SAND SS.pdf", "page_start": 63, "page_end": 68},
    {"file_name": "SALM NO.pdf", "page_start": 20, "page_end": 85},
    {"file_name": "SBRY LN.pdf", "page_start": 13, "page_end": 99},
    {"file_name": "RR LN.pdf", "page_start": 30, "page_end": 101},
    {"file_name": "ROG SW.pdf", "page_start": 84, "page_end": 190},
    {"file_name": "RO SW.pdf", "page_start": 84, "page_end": 190},
    {"file_name": "RBLX US.pdf", "page_start": 21, "page_end": 28},
    {"file_name": "REH AU.pdf", "page_start": 12, "page_end": 25},
    {"file_name": "REC IM.pdf", "page_start": 24, "page_end": 47},
    {"file_name": "RKT LN.pdf", "page_start": 16, "page_end": 160},
    {"file_name": "RAND NA.pdf", "page_start": 131, "page_end": 158},
    {"file_name": "PUM GR.pdf", "page_start": 14, "page_end": 144},
    {"file_name": "PUB FP.pdf", "page_start": 49, "page_end": 212},
    {"file_name": "PRX NA.pdf", "page_start": 17, "page_end": 50},
    {"file_name": "PRX NA.pdf", "page_start": 92, "page_end": 145},
    {"file_name": "PAH3 GR.pdf", "page_start": 60, "page_end": 110},
    {"file_name": "PINS US.pdf", "page_start": 6, "page_end": 49},
    {"file_name": "PLS AU.pdf", "page_start": 12, "page_end": 41},
    {"file_name": "PSON LN.pdf", "page_start": 17, "page_end": 63},
    {"file_name": "PANW US.pdf", "page_start": 17, "page_end": 26},
    {"file_name": "ORK NO.pdf", "page_start": 21, "page_end": 170},
    {"file_name": "ORI AU.pdf", "page_start": 40, "page_end": 124},
    {"file_name": "3288 JP.pdf", "page_start": 15, "page_end": 22},
    {"file_name": "OCI NA.pdf", "page_start": 36, "page_end": 77},
    {"file_name": "OCI NA.pdf", "page_start": 217, "page_end": 239},
    {"file_name": "OCDO LN.pdf", "page_start": 10, "page_end": 32},
    {"file_name": "NZYMB DC.pdf", "page_start": 19, "page_end": 64},
    {"file_name": "NZYMB DC.pdf", "page_start": 73, "page_end": 76},
    {"file_name": "NOVOB DC.pdf", "page_start": 11, "page_end": 32},
    {"file_name": "NOVOB DC.pdf", "page_start": 89, "page_end": 97},
    {"file_name": "NHY NO.pdf", "page_start": 78, "page_end": 122},
    {"file_name": "NDA SS.pdf", "page_start": 316, "page_end": 359},
    {"file_name": "8604 JP.pdf", "page_start": 17, "page_end": 18},
    {"file_name": "8604 JP.pdf", "page_start": 27, "page_end": 42},
    {"file_name": "8604 JP.pdf", "page_start": 48, "page_end": 49},
    {"file_name": "NN NA.pdf", "page_start": 32, "page_end": 140},
    {"file_name": "NIBEB SS.pdf", "page_start": 150, "page_end": 178},
    {"file_name": "NESTE FH.pdf", "page_start": 24, "page_end": 146},
    {"file_name": "NA CN.pdf", "page_start": 10, "page_end": 12},
    {"file_name": "NAB AU.pdf", "page_start": 19, "page_end": 80},
    {"file_name": "9104 JP.pdf", "page_start": 21, "page_end": 37},
    {"file_name": "4188 JP.pdf", "page_start": 49, "page_end": 94},
    {"file_name": "9962 JP.pdf", "page_start": 11, "page_end": 22},
    {"file_name": "MCY NZ.pdf", "page_start": 17, "page_end": 30},
    {"file_name": "MRK GR.pdf", "page_start": 111, "page_end": 157},
    {"file_name": "MPACT SP.pdf", "page_start": 94, "page_end": 140},
    {"file_name": "MLT SP.pdf", "page_start": 110, "page_end": 161},
    {"file_name": "6586 JP.pdf", "page_start": 9, "page_end": 25},
    {"file_name": "LONN SW.pdf", "page_start": 22, "page_end": 35},
    {"file_name": "LIFCOB SS.pdf", "page_start": 14, "page_end": 29},
    {"file_name": "FWONK US.pdf", "page_start": 22, "page_end": 24},
    {"file_name": "LSXMK US.pdf", "page_start": 22, "page_end": 24},
    {"file_name": "LSXMA US.pdf", "page_start": 22, "page_end": 24},
    {"file_name": "LBRDK US.pdf", "page_start": 22, "page_end": 24},
    {"file_name": "LLC AU.pdf", "page_start": 32, "page_end": 106},
    {"file_name": "4151 JP.pdf", "page_start": 20, "page_end": 59},
    {"file_name": "PHIA NA.pdf", "page_start": 45, "page_end": 86},
    {"file_name": "PHIA NA.pdf", "page_start": 266, "page_end": 296},
    {"file_name": "AD NA.pdf", "page_start": 101, "page_end": 142},
    {"file_name": "KOG NO.pdf", "page_start": 33, "page_end": 96},
    {"file_name": "9766 JP.pdf", "page_start": 6, "page_end": 7},
    {"file_name": "6861 JP.pdf", "page_start": 23, "page_end": 37},
    {"file_name": "JMAT LN.pdf", "page_start": 34, "page_end": 69},
    {"file_name": "JMAT LN.pdf", "page_start": 83, "page_end": 136},
    {"file_name": "JMT PL.pdf", "page_start": 160, "page_end": 336},
    {"file_name": "JD LN.pdf", "page_start": 53, "page_end": 91},
    {"file_name": "8001 JP.pdf", "page_start": 39, "page_end": 54},
    {"file_name": "INVEB SS.pdf", "page_start": 15, "page_end": 22},
    {"file_name": "INVEB SS.pdf", "page_start": 60, "page_end": 80},
    {"file_name": "INVEA SS.pdf", "page_start": 15, "page_end": 22},
    {"file_name": "INVEA SS.pdf", "page_start": 60, "page_end": 80},
    {"file_name": "INGA NA.pdf", "page_start": 20, "page_end": 102},
    {"file_name": "INDT SS.pdf", "page_start": 7, "page_end": 56},
    {"file_name": "IMB LN.pdf", "page_start": 35, "page_end": 65},
    {"file_name": "IMB LN.pdf", "page_start": 94, "page_end": 155},
    {"file_name": "IEL AU.pdf", "page_start": 13, "page_end": 17},
    {"file_name": "HSBA LN.pdf", "page_start": 42, "page_end": 97},
    {"file_name": "HOLMB SS.pdf", "page_start": 16, "page_end": 53},
    {"file_name": "HEIA NA.pdf", "page_start": 126, "page_end": 183},
    {"file_name": "HEIO NA.pdf", "page_start": 126, "page_end": 183},
    {"file_name": "HEI GR.pdf", "page_start": 21, "page_end": 64},
    {"file_name": "HL LN.pdf", "page_start": 26, "page_end": 44},
    {"file_name": "HL LN.pdf", "page_start": 61, "page_end": 123},
    {"file_name": "HLMA LN.pdf", "page_start": 44, "page_end": 97},
    {"file_name": "HLMA LN.pdf", "page_start": 106, "page_end": 159},
    {"file_name": "HAL US.pdf", "page_start": 14, "page_end": 72},
    {"file_name": "HLN LN.pdf", "page_start": 20, "page_end": 55},
    {"file_name": "HLN LN.pdf", "page_start": 64, "page_end": 107},
    {"file_name": "GBLB BB.pdf", "page_start": 138, "page_end": 183},
    {"file_name": "GJF NO.pdf", "page_start": 40, "page_end": 131},
    {"file_name": "GETIB SS.pdf", "page_start": 34, "page_end": 74},
    {"file_name": "GENS SP.pdf", "page_start": 20, "page_end": 47},
    {"file_name": "6504 JP.pdf", "page_start": 9, "page_end": 29},
    {"file_name": "FRE GR.pdf", "page_start": 101, "page_end": 285},
    {"file_name": "FLTR LN.pdf", "page_start": 38, "page_end": 189},
    {"file_name": "FPH NZ.pdf", "page_start": 34, "page_end": 97},
    {"file_name": "FER SM.pdf", "page_start": 26, "page_end": 111},
    {"file_name": "FER SM.pdf", "page_start": 120, "page_end": 133},
    {"file_name": "BALDB SS.pdf", "page_start": 25, "page_end": 40},
    {"file_name": "BALDB SS.pdf", "page_start": 94, "page_end": 116},
    {"file_name": "EVO SS.pdf", "page_start": 31, "page_end": 74},
    {"file_name": "ESSITYB SS.pdf", "page_start": 50, "page_end": 71},
    {"file_name": "EBS AV.pdf", "page_start": 49, "page_end": 126},
    {"file_name": "EQT SS.pdf", "page_start": 28, "page_end": 50},
    {"file_name": "EQT SS.pdf", "page_start": 68, "page_end": 97},
    {"file_name": "EPIB SS.pdf", "page_start": 34, "page_end": 53},
    {"file_name": "EPIB SS.pdf", "page_start": 144, "page_end": 161},
    {"file_name": "EPIA SS.pdf", "page_start": 34, "page_end": 53},
    {"file_name": "EPIA SS.pdf", "page_start": 144, "page_end": 161},
    {"file_name": "ENG SM.pdf", "page_start": 49, "page_end": 160},
    {"file_name": "EMBRACB SS.pdf", "page_start": 42, "page_end": 93},
    {"file_name": "DB1 GR.pdf", "page_start": 49, "page_end": 116},
    {"file_name": "DB1 GR.pdf", "page_start": 128, "page_end": 132},
    {"file_name": "DSY FP.pdf", "page_start": 36, "page_end": 106},
    {"file_name": "AM FP.pdf", "page_start": 28, "page_end": 37},
    {"file_name": "CSL AU.pdf", "page_start": 37, "page_end": 62},
    {"file_name": "SGO FP.pdf", "page_start": 22, "page_end": 52},
    {"file_name": "CBA AU.pdf", "page_start": 10, "page_end": 115},
    {"file_name": "BRBY LN.pdf", "page_start": 52, "page_end": 143},
    {"file_name": "BRBY LN.pdf", "page_start": 152, "page_end": 219},
    {"file_name": "BRO US.pdf", "page_start": 12, "page_end": 18},
    {"file_name": "BHP AU.pdf", "page_start": 28, "page_end": 64},
    {"file_name": "BAS GR.pdf", "page_start": 100, "page_end": 194},
    {"file_name": "BARC LN.pdf", "page_start": 26, "page_end": 44},
    {"file_name": "BARC LN.pdf", "page_start": 63, "page_end": 65},
    {"file_name": "BARC LN.pdf", "page_start": 69, "page_end": 263},
    {"file_name": "BAC US.pdf", "page_start": 36, "page_end": 57},
    {"file_name": "BALN SW.pdf", "page_start": 50, "page_end": 101},
    {"file_name": "BA LN.pdf", "page_start": 38, "page_end": 79},
    {"file_name": "BANB SW.pdf", "page_start": 20, "page_end": 42},
    {"file_name": "6845 JP.pdf", "page_start": 16, "page_end": 51},
    {"file_name": "AIA NZ.pdf", "page_start": 10, "page_end": 37},
    {"file_name": "ATCOB SS.pdf", "page_start": 5, "page_end": 12},
    {"file_name": "ATCOB SS.pdf", "page_start": 33, "page_end": 43},
    {"file_name": "ATCOB SS.pdf", "page_start": 126, "page_end": 145},
    {"file_name": "ATCOA SS.pdf", "page_start": 5, "page_end": 12},
    {"file_name": "ATCOA SS.pdf", "page_start": 33, "page_end": 43},
    {"file_name": "ATCOA SS.pdf", "page_start": 126, "page_end": 145},
    {"file_name": "ASX AU.pdf", "page_start": 24, "page_end": 72},
    {"file_name": "AKE FP.pdf", "page_start": 16, "page_end": 69},
    {"file_name": "9202 JP.pdf", "page_start": 22, "page_end": 52},
    {"file_name": "ALO FP.pdf", "page_start": 181, "page_end": 342},
    {"file_name": "AKZA NA.pdf", "page_start": 29, "page_end": 97},
    {"file_name": "AGS BB.pdf", "page_start": 26, "page_end": 85},
    {"file_name": "ADYEN NA.pdf", "page_start": 14, "page_end": 81},
    {"file_name": "ADS GR.pdf", "page_start": 16, "page_end": 64},
    {"file_name": "ADEN SW.pdf", "page_start": 14, "page_end": 41},
    {"file_name": "ADEN SW.pdf", "page_start": 59, "page_end": 101},
    {"file_name": "ACS SM.pdf", "page_start": 86, "page_end": 213},
]

OUTPUT_EXCEL_FILE_NAME = "!output.xlsx"
LOG_FILE_NAME = "!logs.txt"

def extract_pdf(file, file_params):

    reader = PdfReader(file)
    extracted_pages = []

    page_numbers_to_include = []
    for item in file_params:
        page_start = item["page_start"]-1
        page_end = item["page_end"]-1
        page_numbers_to_include.extend(range(page_start, page_end+1))

    if not file_params:
        for page in reader.pages:
            page_data = page.extract_text()
            extracted_pages.append(page_data)
    else:
        for index, page in enumerate(reader.pages):
            if index in page_numbers_to_include:
                page_data = page.extract_text()
                extracted_pages.append(page_data)
            else:
                continue

    tokenized_page_list = []
    for page in extracted_pages:
        tokenized_page = tokenize.sent_tokenize(page)
        tokenized_page_list.append(tokenized_page)


    for index, page in enumerate(tokenized_page_list):
        tokenized_page_list[index] = [x.replace('\n',' ').strip() for x in page]

    file_sentences_list = []
    for page in tokenized_page_list:
        for sentence in page:
            file_sentences_list.append(sentence)

    return file_sentences_list

def if_string_has_number(input_string):
    return any(char.isdigit() for char in input_string)

def if_string_has_currency(input_string):

    input_string =  re.sub('([+-]?(?=\.\d|\d)(?:\d+)?(?:\.?\d*))(?:[Ee]([+-]?\d+))?', r' \1 ', input_string)

    currency_found = False
    found_currency_data = list(lexnlp.extract.en.money.get_money(input_string))

    if len(found_currency_data) != 0:
        currency_found = True

    return currency_found

if __name__ == '__main__':

    with open(LOG_FILE_NAME, 'w', encoding="UTF-8") as f:
        f.write("PDF Scanner logs\n\n")

    if os.path.exists(OUTPUT_EXCEL_FILE_NAME):
        try:
            os.rename(OUTPUT_EXCEL_FILE_NAME, OUTPUT_EXCEL_FILE_NAME)
        except OSError as e:
            print('Access-error on file "' + OUTPUT_EXCEL_FILE_NAME + '"! \n' + str(e))
            raise Exception(f"Close {OUTPUT_EXCEL_FILE_NAME} and try again")

    #Create list of .pdf files in folder where python script is located
    files = [file for file in os.listdir('.') if os.path.isfile(file)]
    files = [file for file in files if file.endswith(".pdf")]

    output_list = []
    log_list = []
    error_counter = 0

    print(f"Job Starting: {len(files)} PDF files found\n")
    for file in files:

        file_params = []
        for file_to_find in PDF_PAGES_TO_INCLUDE:
            if file_to_find["file_name"] == file:
                file_params.append(file_to_find)


        log_string = f"Scanning {file} ... "
        print(log_string, end="")

        try:
            extracted_sentences = extract_pdf(file, file_params)
            sentences_count = len(extracted_sentences)

            sentences_list_with_ratings = []

            for sentence in extracted_sentences:

                item = {"E": None, "S": None, "G": None, "sentence":sentence}
                sentences_list_with_ratings.append(item)

            for item in sentences_list_with_ratings:
                sentence = item["sentence"]

                keywords_environmental_in_sentence = False
                keywords_social_in_sentence = False
                keywords_governance_in_sentence = False

                for keyword in ENVIROMENTAL_KEYWORDS:
                    if re.search(rf"\b{keyword.lower()}\b", sentence.lower()):
                        keywords_environmental_in_sentence = True
                        break

                for keyword in SOCIAL_KEYWORDS:
                    if re.search(rf"\b{keyword.lower()}\b", sentence.lower()):
                        keywords_social_in_sentence = True
                        break

                for keyword in GOVERNANCE_KEYWORDS:
                    if re.search(rf"\b{keyword.lower()}\b", sentence.lower()):
                        keywords_governance_in_sentence = True
                        break

                if keywords_environmental_in_sentence == False:
                    item["E"] = 0

                if keywords_social_in_sentence == False:
                    item["S"] = 0

                if keywords_governance_in_sentence == False:
                    item["G"] = 0


                #if keyword accurance in sentance occured:
                string_has_number = False
                string_has_currency = False

                string_has_number = if_string_has_number(sentence)
                string_has_currency = if_string_has_currency(sentence)

                #tweak item rating number based on has number or has currency
                if keywords_environmental_in_sentence == True:
                    item["E"] = 1

                    if string_has_number:
                        item["E"] = 2

                    if string_has_number and string_has_currency:
                        item["E"] = 3

                if keywords_social_in_sentence == True:
                    item["S"] = 1

                    if string_has_number:
                        item["S"] = 2

                    if string_has_number and string_has_currency:
                        item["S"] = 3

                if  keywords_governance_in_sentence == True:
                    item["G"] = 1

                    if string_has_number:
                        item["G"] = 2

                    if string_has_number and string_has_currency:
                        item["G"] = 3

            e0_total = 0
            e1_total = 0
            e2_total = 0
            e3_total = 0

            s0_total = 0
            s1_total = 0
            s2_total = 0
            s3_total = 0

            g0_total = 0
            g1_total = 0
            g2_total = 0
            g3_total = 0

            for sentence in sentences_list_with_ratings:

                e0_total += 1 if sentence["E"] == 0 else 0
                e1_total += 1 if sentence["E"] == 1 else 0
                e2_total += 1 if sentence["E"] == 2 else 0
                e3_total += 1 if sentence["E"] == 3 else 0

                s0_total += 1 if sentence["S"] == 0 else 0
                s1_total += 1 if sentence["S"] == 1 else 0
                s2_total += 1 if sentence["S"] == 2 else 0
                s3_total += 1 if sentence["S"] == 3 else 0

                g0_total += 1 if sentence["G"] == 0 else 0
                g1_total += 1 if sentence["G"] == 1 else 0
                g2_total += 1 if sentence["G"] == 2 else 0
                g3_total += 1 if sentence["G"] == 3 else 0

            output_row = {
                "name_of_file": os.path.splitext(file)[0],
                "total_sentences": sentences_count,
                "E0": e0_total,
                "E1": e1_total,
                "E2": e2_total,
                "E3": e3_total,
                "S0": s0_total,
                "S1": s1_total,
                "S2": s2_total,
                "S3": s3_total,
                "G0": g0_total,
                "G1": g1_total,
                "G2": g2_total,
                "G3": g3_total
                }

            output_list.append(output_row)

            log_string += "success"
            print("success")

            with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
                f.write(log_string + "\n")
        except:
            log_string += "error"
            print("error")

            error_counter += 1

            with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
                f.write(log_string + "\n")

    #end of files loop
    try:
        df = pd.DataFrame(output_list, columns=["name_of_file", "total_sentences",
                                                "E0","E1","E2","E3",
                                                "S0","S1","S2","S3",
                                                "G0","G1","G2","G3"])

        df.to_excel(OUTPUT_EXCEL_FILE_NAME, index=False)
    except:

        with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
            f.write("Error creating result excel file" + "\n")

        print("Error creating result excel file")
        raise Exception("Error creating result excel file")

    with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
        f.write(f"\nJob finished: files processed={len(files)}, errors={error_counter} (ctrl+f search for 'error')")

    print(f"\nJob finished: files processed={len(files)}, errors={error_counter}")




