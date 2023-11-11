SELECT MAXIMO_TICKET.TICKETID, MAXIMO_TICKET.CLASS, MAXIMO_TICKET.DESCRIPTION, MAXIMO_TICKET.STATUS, MAXIMO_TICKET.TARGETFINISH, MAXIMO_TICKET.SC_ANALYSISTYPE AS A_TYPE, MAXIMO_TICKET.SC_ANALDUEDATE AS A_DUE, MAXIMO_TICKET.SC_RESPMGR AS MCR, MAXIMO_TICKET.SC_OWNER AS OWNER, MAXIMO_TICKET.SC_AITYPE
FROM MAXIMO_TICKET
WHERE (
     (
     (MAXIMO_TICKET.STATUS) Not ALike "%CLOSE%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%CANCEL%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%ACTCOMP%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%APPRMGMT%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%INPROGSR%" 
     ) 
     AND 
    (
     (MAXIMO_TICKET.SC_OWNER) 
           In 
           (
         "VBALAKRI",
          "EJBEAN",
          "BHBENDER",
          "X2CTBLAC",
          "X2JBLAZE",
          "AMBOERNE",
          "MJBOUCHE",
          "PGBRADAT",
          "X2JRBRAN",
          "X2JGBRIN",
          "JTCORBET",
          "X2BCUSIC",
          "ddarr",
          "X2ELDICK",
          "X2NERTLE",
          "X2MFRENC",
          "pgibbs",
          "X2LGIHYE",
          "X2NGRAFT",
          "X2RGRIME",
          "X2DEHARM",
          "X2BHIRMA",
          "CMHOWARD",
          "AHUSSEIN",
          "X2LKATZE",
          "X2PGMANS",
          "DCMCCORM",
          "X2RHNORV",
          "DAOGLESB",
          "X2BDORTI",
          "RFPILUSO",
          "PDPOTTER",
          "RLRUNYON",
          "X2MSSHOO",
          "DESPIELM",
          "X2DSSPIN",
          "X2DSREDN",
          "X2KPVIKA",
          "X2BWERNE",
          "X2TAWOOL",
          "KAYENNER"
     )
     )
) 
OR 
(
     (
     (MAXIMO_TICKET.STATUS) Not ALike "%CLOSE%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%CANCEL%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%ACTCOMP%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%APPRMGMT%" 
     And (MAXIMO_TICKET.STATUS) Not ALike "%INPROGSR%"
     ) 
     AND 
     (
     (MAXIMO_TICKET.TARGETFINISH)<Now()+21
     ) 
     AND 
     (
     (MAXIMO_TICKET.SC_RESPMGR)="I&CCONST" 
     Or (MAXIMO_TICKET.SC_RESPMGR)="DIGSYSMGR"
     )
)
ORDER BY MAXIMO_TICKET.CLASS, MAXIMO_TICKET.TARGETFINISH, MAXIMO_TICKET.SC_ANALDUEDATE;
