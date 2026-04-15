export const RULES = {
  versionLabel: "Final Project Grader 2025 v3.0",
  matlabVersion: "Final_project_autograde_Olmstead_Fall_2025_v02",
  scoringReference:
    "Final Project web scoring must match Final_project_autograde_Olmstead_Fall_2025_v02.m exactly. Do not copy Final Exam values or point maps into this repo.",
  sheets: {
    Aero: {
      fallbackCells: ["G3", "G4", "G10", "G11", "A15", "A16"],
    },
    Miss: {
      fallbackCells: [
        "C48", "D48", "E48", "F48", "G48", "H48", "I48", "J48", "K48", "L48", "M48", "N48",
        "C49", "D49", "E49", "F49", "G49", "H49", "I49", "J49", "K49", "L49", "M49", "N49",
      ],
    },
    Main: {
      fallbackCells: [
        "S3", "T3", "U3", "V3", "W3", "X3", "Y3",
        "S4", "T4", "U4", "V4", "W4", "X4", "Y4",
        "S5", "S6", "S7", "S8", "S9",
        "T6", "U6", "V6", "W6", "X6", "Y6",
        "T7", "U7", "V7", "W7", "X7", "Y7",
        "T8", "U8", "V8", "W8", "X8", "Y8",
        "T9", "U9", "V9", "W9", "X9", "Y9",
        "S12", "S13",
        "AB3", "AB4", "X12", "X13", "Y37",
        "M10", "O10", "P10", "Q10",
        "O18", "X40", "Q23", "Q31", "N31",
        "P13", "Q13",
        "B19", "C19", "D19", "H19",
        "B21", "C21", "D21", "H21",
        "B23", "C23", "D23", "H23",
        "C24", "D24", "H24",
        "C26", "D26", "H26",
        "B27", "C27", "D27", "H27",
        "F31", "F32", "H29", "I29", "O15",
        "B34", "B35", "B36", "B37", "B38", "B39", "B40", "B41", "B42", "B43",
        "B44", "B45", "B46", "B47", "B48", "B49", "B50", "B51", "B52", "B53",
        "E34", "E35", "E36", "E37", "E38", "E39", "E40", "E41", "E42", "E43",
        "E44", "E45", "E46", "E47", "E48", "E49", "E50", "E51", "E52", "E53",
        "K33", "L33", "M33", "N33", "P33", "R33", "S33", "V33", "W33",
        "K35", "L35", "M35", "N35", "P35", "R35", "S35", "V35", "W35",
        "K36", "L36", "M36", "N36", "P36", "R36", "S36", "V36", "W36",
        "K38", "L38", "M38", "N38", "P38", "R38", "S38", "V38", "W38",
        "K39", "L39", "M39", "N39", "P39", "R39", "S39", "V39", "W39",
        "B32", "C18", "C19", "C20", "C21", "C22", "C23", "C24", "C25", "C26", "C27",
        "D18", "D19", "D20", "D21", "D22", "D23", "D24", "D25", "D26", "D27",
        "H18", "H19", "H20", "H21", "H22", "H23", "H24", "H25", "H26", "H27",
        "D18", "D23", "D52", "F52",
        "H24", "E52",
      ],
    },
    Consts: {
      fallbackCells: ["K22", "K23", "K24", "K26", "K27", "K28", "K29", "K32", "K33", "L33", "AO42", "AQ41"],
    },
    Gear: {
      fallbackCells: ["J20", "L20", "L21", "M20", "M21", "N20", "N21"],
    },
    Geom: {
      fallbackCells: ["C8", "C10", "M152", "K15", "L155", "L38", "L40", "L41", "N44", "L117", "L118", "N121"],
    },
  },
  macroWarningText:
    "This workbook was saved as .xlsx; save as a macro-enabled (.xlsm) file outside Downloads to retain JET macros and Ps plotting.",
};
