
# Dictionary for each file
class WorklistDictionary:

    """
        Full Definition of Each report
    """



    def __init__(self):

        self.Scott : dict[str:str] = {
            "aaall": "METALS",
            "ext-nd": "ORGPREP",
            "gcs": "GCSEMI",
            "gcv": "GCVOA",
            "gnallnd": "GENCHEM",
            "hgall": "HG",
            "mss": "MSSEMI",
            "msvoa": "MSVOA"
        }

        self.Wheat_Ridge: dict[str:str] = {
            "aaallnd": "METALS",
            "gcs": "GCSEMI",
            "gcvoa": "GCVOA",
            "gnall": "GENCHEM",
            "hgall": "HG",
            "mss": "MSSEMI",
            "msvoa": "MSVOA",
            "orgprep": "ORGPREP"

        }

        self.Orlando: dict[str:str] ={
            "gcvnd":"GCVOA",
            "gcnd":"GCSEMI",
            "extnd":"ORGPREP",
            "gcmsnd":"MSSEMI",
            "msvnd": "MSVOA",
            "aaallnd": "METALS",
            "gnallnd": "GENCHEM",
            "hg":"HG",
            "lc-qqq": "LCMSPFAS",
            "lc-qqqprep": "EXTLCMS"
        }


        self.Dayton: dict[str:str] = {
            "mss": "MSSEMI", 
            "gcs": "GCSEMI",
            "msvoa": "MSVOA",
            "gcvoa": "GCVOA",
            "msair" : "MSAIR",
            "gcair": "GCAIR",
            "gnallnd": "GENCHEM",
            "aaallnd": "METALS",
            "ext-nd": "ORGPREP",
            "lcmspfas": "LCMSPFAS",
            "hgall": "HG",
            "extlcms": "EXTLCMS"
        }


        self.Dayton_Late = {
            "raviv": "RAVIV",
            "langan": "LANGAN_SBMT",
            "ttnjp": "TTNJP",
            "ttscji": "TTSCJI",
            "arecnjh": "ARECNJH",
            "alms": "ALMS"
        }

        self.Wheat_Ridge_Late = {
            "aecomhih": "AECOMHIH",
            "chevrcog": "CHEVRCOG"
        }

        self.Orlando_Late = {
            "aecomcod":"AECOMCOD",
            "aecomtxa": "AECOMTXA",
            "aerotnor":"AEROTNOR",
            "plexvavb": "PLEXVAVB"
        }
