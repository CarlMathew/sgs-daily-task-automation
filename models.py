import datetime
from sqlalchemy import Column, String, Integer, ForeignKey, DateTime, create_engine
from sqlalchemy.orm import declarative_base
import sqlite3
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine



 


Base = declarative_base()



def create_database_sqlite(name: str) -> None:
    """
        Create a database for nam, and order to store all of the data inside.

        Args:
            name: str -> name of the database
    """
    if name:
        conn = sqlite3.connect(name)
        print(f"Database Created: {name}")
        conn.close()
    else:
        raise ValueError("Please Provide a name")


def database_engine(url:str):  
    """
        Connect db into your python

        Return: engine 
    """
    engine = create_engine(url=url)
    return engine


class NAM_Information_Table(Base):

        """Information on everysite"""
        __tablename__ = "NAM_Information_Tbl"

        id = Column(Integer, primary_key=True)
        site = Column(String(200), nullable = False)
        abbreviate_site = Column(String(200), nullable = False)
        anita_hostname= Column(String(200), nullable= False) 

        def __repr__(self):
            return f"Added new site: {self.site}"
    

    
class Scott_Late_Count(Base):
        """ Table that consite the data of late report """
        __tablename__ = "NAM_Report_Count_Tbl"

        id = Column(Integer, primary_key=True)
        worklist_name = Column(String(200))
        late_count = Column(Integer)
        total_count = Column(Integer, nullable=True)
        element_ref_excel = Column(String(200))
        site_id = Column(Integer, ForeignKey("NAM_Information_Tbl.id"))
        data_collected = Column(DateTime, default=datetime.datetime.now)



def insert_site_into_database(database_name: str) -> None:
    """
        Insert all the information of the site into the database.

        Args:
            database_name: str -> name of the database
        
    """
    sites = [
        {
            "site": "Scott",
            "abbreviate_site": "la",
            "anita_hostname": "use-idb062"
         },

         {
            "site": "Wheat Ridge",
            "abbreviate_site": "co",
            "anita_hostname": "use-idb057"   
         },
         {
             "site": "Orlando",
             "abbreviate_site": "fl",
             "anita_hostname": "use-idb066"
         },
         {
             "site": "Dayton", 
             "abbreviate_site": "nj",
             "anita_hostname": "use-idb059"

         }
    ]

    engine = database_engine(f"sqlite:///{database_name}")
    Base.metadata.create_all(engine)  

    Session = sessionmaker(bind=engine)
    session = Session()

    try:

        session.add_all([NAM_Information_Table(**site) for site in sites])
        session.commit()
        print("Successfully Added the site")
    finally:
        session.close()
        print("Error on Adding the site")
    

if __name__ == "__main__": 
       
    try:

        
        name_of_database: str = "database_file/nam_database.db"

    #Start operation for creating database and insert data into database
        create_database_sqlite(name_of_database)

        engine = create_engine(url="sqlite:///database_file/nam_database.db")
        Base.metadata.create_all(bind = engine)

        insert_site_into_database(name_of_database)

        print("Operation for database successfully.")
    except Exception as e:
        print(f"Error Create table: {e}")
