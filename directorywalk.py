import os
import zipfile
import xmltodict
import imghdr
import sqlite3
import time, datetime
import logging
import configparser
import io
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
## Stuff
config = configparser.ConfigParser()

config.read('twirler.conf')


def dbInit(conn):
    
    
    c = conn.cursor()
    
    results = c.execute(""" select exists 
                            (select * from sqlite_master 
                            where type='table' and name = 'comic');
                        """).fetchone()[0]
    
    if results != 1:
        
        logging.info('No comic table in database.  Creating one.')
        
        c.execute("""
            create table comic
            (
                CVDB int,
                Series text,
                Volume text,
                Number int,
                Title text,
                AlternateSeries text,
                Summary,
                Year int,
                Month int,
                Writer text,
                Penciller text,
                Inker text,
                Colorist text,
                Letterer text,
                CoverArtist,
                Editor text,
                Publisher text,
                Imprint text,
                DateAdded text,
                URL text,
                Path text,
                PRIMARY KEY (CVDB),
                UNIQUE (Series, Volume, Number)
            )
            """)
    conn.commit()
    c.close()
                
def findCVDB(info):

    # initialise the cvdb identifier as None
    cvdb=None
    
    # Let's look at the web url to see if it is from comicvine
    # If so, we should be able to extract the id from the url
    url = info['ComicInfo'].get('Web')

    if url is not None:
        if 'comicvine' in url and url.startswith("http"):
            # Looks like a comicvine url 
            start = url.rfind('-')+1
            if url.endswith("/"):
                end = len(url)-1
            else:
                end = len(url)
            
            if start != -1:
                cvdb = url[start:end]
    
    return cvdb
    
    
def updateComics(conn, scanDir):
    
    logging.info("Looking for new comics in "+scanDir)
    
    for root, dirs, files in os.walk(scanDir, topdown=False):
        
        CurrentDate = datetime.datetime.today()
        
        for name in files:
            path = os.path.join(root, name)
            
            if zipfile.is_zipfile(path):
                archive = zipfile.ZipFile(path, 'r')
                if 'ComicInfo.xml' in archive.namelist():
                    info = xmltodict.parse(archive.read('ComicInfo.xml'))
                    
                    cvdb = findCVDB(info)
                    if cvdb is None:
                        logging.warning("""Could not find CVDB identifier for 
                                        """+info['ComicInfo'].get('Series')+"""
                                        , Issue """+info['ComicInfo'].get('Number')+"""
                                        . Consider trying to rescrape the metadata.""")
                    else:
                        row = ( cvdb,
                                info['ComicInfo'].get('Series'),
                                info['ComicInfo'].get('Volume'),
                                info['ComicInfo'].get('Number'),
                                info['ComicInfo'].get('Title'),
                                info['ComicInfo'].get('AlternateSeries'),
                                info['ComicInfo'].get('Summary'),
                                info['ComicInfo'].get('Year'),
                                info['ComicInfo'].get('Month'),
                                info['ComicInfo'].get('Writer'),
                                info['ComicInfo'].get('Penciller'),
                                info['ComicInfo'].get('Inker'),
                                info['ComicInfo'].get('Colorist'),
                                info['ComicInfo'].get('Letterer'),
                                info['ComicInfo'].get('CoverArtist'),
                                info['ComicInfo'].get('Editor'),
                                info['ComicInfo'].get('Publisher'),
                                info['ComicInfo'].get('Imprint'),
                                info['ComicInfo'].get('Web'),
                                path,
                                cvdb,
                                CurrentDate
                            )
                        
                        c = conn.cursor()
                        
                        c.execute("""
                                insert or replace into comic
                                (CVDB, Series, Volume, Number, Title, AlternateSeries, Summary,
                                Year, Month, Writer, Penciller, inker, Colorist, Letterer,
                                CoverArtist, Editor, Publisher, Imprint, URL, Path, DateAdded)
                                values  (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,
                                (coalesce((select DateAdded from comic where cvdb=?),?)))
                                """, row)
                                                    
                    conn.commit()
                    c.close()

def prepareHTML(conn, scanDate):
    
    logging.info('Preparing HTML with all updates since '+scanDate)
    
    header='''
            <head>
                <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
            </head>
            '''
    
    newSeries = ''
    newComics = getComics(conn, scanDate)
    
    body='''
        <body style='''+config.get('Styles', 'body')+'''>
            <h1 style='''+config.get('Styles', 'h1')+'''>
                New Comics
            </h1>
            <p style='''+config.get('Styles', 'p')+'''>
                Here are the latest new comics downloaded, since '''+str(scanDate)+'''
            </p>
            '''+newSeries+newComics+'''
        </body>'''
    
    footer='''<hr style='''+config.get('Styles', 'hr')+'''/>
            <small>
                This email was bought to you by Twirler
            </small>
            '''
    
    ## Build complete document
    content='<!DOCTYPE html><html lang="en">'+header+body+footer+'</html>'
    
    return content


def getComics(conn, scanDate):
    
    conn.row_factory = sqlite3.Row
    
    c = conn.cursor()
    
    c.execute("""
            select  * 
            from    comic 
            where   DateAdded > ? 
            order   by Series, Volume, Number
        """, (scanDate,))
    comics = c.fetchall()    
    
    if len(comics) > 0:
        
        logging.info('Found '+str(len(comics))+ ' new comics.')
        
        prevSeries=""
        heading='''
                <hr style='''+config.get('Styles', 'hr')+'''/>
                <h2 style='''+config.get('Styles', 'h2')+'''>
                    New Comics
                </h2>
                '''
        body = ""
        
        for row in comics:
            if row['Series'] != prevSeries:
                seriesTitle='''
                    <h3 style="'''+config.get('Styles', 'h3')+'''">
                        '''+row['Series']+''' ('''+row['Volume']+''')
                    </h3>
                    '''
                if row['Publisher'] is not None:
                    if len(row['Publisher']) != 0:
                        seriesTitle+='''
                            <p style="'''+config.get('Styles', 'publisher')+'''">
                                '''+row['Publisher']+'''
                            </p>
                        '''
            else:
                seriesTitle=''
            
            issueTitle = '''
                <a style="'''+config.get('Styles', 'a')+'''" href="'''+row['URL']+'''">
                    <h4 style="'''+config.get('Styles', 'h4')+'''">
                        Issue #'''+str(row['Number'])+'''
                    </h4>
                </a>
                '''
            
            if row['Summary'] is None:
                summaryText='No story details available.'
            else:
                summaryText=smartTrunc(row['Summary'])
                
            if summaryText.endswith('...'):
                summaryText = summaryText + '''
                    <a href="'''+row['URL']+'''" style='''+config.get('Styles', 'details')+'''>[more]</a>'''
            
            summary = '''
                <p style='''+config.get('Styles', 'plot')+'''>
                    '''+summaryText+'''
                </p>'''
            
            
            d=[]
            if row['Writer'] is not None:
                if len(row['Writer']) != 0:
                    d.append('<b>Writer: </b>'+row['Writer'])
            
            if row['Penciller'] is not None:
                if len(row['Penciller']) != 0:
                    d.append('<b>Penciller: </b>'+row['Penciller'])
            
            if row['Inker'] is not None:
                if len(row['Inker']) != 0:
                    d.append('<b>Inker: </b>'+row['Inker'])
            
            if row['Colorist'] is not None:
                if len(row['Colorist']) != 0:
                    d.append('<b>Colorist: </b>'+row['Colorist'])
            
            if row['Letterer'] is not None:
                if len(row['Letterer']) != 0:
                    d.append('<b>Letterer: </b>'+row['Letterer'])
            
            if row['CoverArtist'] is not None:
                if len(row['CoverArtist']) != 0:
                    d.append('<b>Cover Artist: </b>'+row['CoverArtist'])
            
            if row['Editor'] is not None:
                if len(row['Editor']) != 0:
                    d.append('<b>Editor: </b>'+row['Editor'])
            
            details = '''
                <p style='''+config.get('Styles', 'details')+'''>
                    '''+" <br/> ".join(d)+'''
                </p>'''
                
            prevSeries = row['Series']
            
            body+=seriesTitle+issueTitle+summary+details
            
    c.close()
    
    return heading+body

def checkNew(conn, scanDate):
    
    c = conn.cursor()
    
    comicCount = c.execute("""
            select  count(*)
            from    comic 
            where   DateAdded > ? 
        """, (scanDate,)).fetchone()[0]
    
    c.close()
    
    return comicCount
    
def sendEmail(content):
            
    #Get data from config
    
    emailFrom = config.get('Twirler', 'email_from')
    emailTo = config.get('Twirler', 'email_to').split(',')
    emailServer = config.get('Twirler', 'email_server')
    emailPort = config.get('Twirler', 'email_port')
    emailLogin = config.get('Twirler', 'email_login')
    emailPassword = config.get('Twirler', 'email_password')
    
    msg=MIMEMultipart('alternative')
    
    msg['Subject'] = "Your latest downloaded comics"
    msg['From'] = emailFrom
    msg['To'] = ",".join(emailTo)
    
    part=MIMEText(content, 'html', 'utf-8')
    
    msg.attach(part)

    logging.info('Sending email to '+msg['To']+'.')

    server = smtplib.SMTP(emailServer, emailPort)
    server.ehlo()
    server.starttls()
    server.login(emailLogin, emailPassword)
    server.sendmail(emailFrom, emailTo, msg.as_string())
    server.close()
    
def main():
    
    scanDir = config.get('Twirler', 'scan_dir')
    scanDays = int(config.get('Twirler', 'scan_days')) * -1
    currentDate = datetime.datetime.today()
    scanDate = (currentDate+datetime.timedelta(days=scanDays)).strftime('%Y-%m-%d %H:%M:%S')
    print(scanDate)
    
    logging.basicConfig(level=logging.DEBUG)
     
    try:
        
        logging.info('Trying to find existing database.')
        
        d = open('comics.db')
    
    except IOError:
    
        logging.warning('No database found. Creating one.')
    
        firstRun = True
    
    
    else:
    
        logging.info('Found one. I hope it''s correct.')

        firstRun = False
    
        d.close()

    conn = sqlite3.connect('comics.db')
    
    dbInit(conn=conn)
    
    updateComics(conn, scanDir)
    
    newComics = checkNew(conn, scanDate)
    
    if newComics > 0:
        logging.info('Found '+str(newComics)+' new entries.  Preparing HTML update.')
        html=prepareHTML(conn, scanDate)
        
        if config.get('Twirler', 'send_email') == "False":
            logging.info('Email set to False.  No email sent.  HTML document will still be saved.')
        else:
            sendEmail(html)
        
        f = io.open('output.htm', 'w', encoding='utf-8')
        f.write(html)
        f.close()
    
    
def smartTrunc(content, length=500, suffix='...'):
    if len(content) <= length:
        return content
    else:
        return content[:length].rsplit(' ', 1)[0]+suffix

if __name__ == "__main__":
    main()