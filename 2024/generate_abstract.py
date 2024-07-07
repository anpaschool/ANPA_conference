import pandas as pd
import numpy as np
from datetime import timedelta
import os
from os.path import join
#from ics import Calendar,Event 


def load_data(filename, sheet_name=None):
  """
  Load data from the merged dataset: merged_v4.csv
  """
  if filename.endswith('.csv'):
    df = pd.read_csv(filename)
  elif filename.endswith('.xlsx'):
    df = pd.read_excel(filename, sheet_name=sheet_name)
  else:
    print ("File format not supported. Could not load data from  {filename}")
  ## Assign AbstractID
  df['abstract_id'] = ['ANPA2023-N000'+str(x) for x in range(1,len(df)+1)]
  cols_ = df.columns
  cols_needed = ['presenter', 'method', 'affiliation', 'manuscript', 'session', 'title',
                  'abstract', 'invited',  'Start Date', 'Start Time', 'End Date', 'End Time', 'abstract_id']
  if len([c for c in cols_needed if c not in cols_]) >0:
    raise ValueError(f"Some of the required columns are not present in the data input: {filename}")  
  df = df[cols_needed]

  return df

def process_data(df):

  col_mapping = {"invited":"Type", "session":"division"}

  df = df.rename(columns=col_mapping)
  #df  = df.astype(str)
  #df[['coauthor_names','coauthor_affiliations']] = df[['coauthor_names','coauthor_affiliations']].fillna(" ")
  ## I don't need all these breaks. Will drop them.
  df = df.loc[~df['Type'].isnull()]
  print (df.isnull().sum() )

  ## Now creating date_time columns.
  df['Date/Time']= df['Start Date'].astype(str)+' '+df['Start Time'].astype(str)

  ## Everything looks OK. Now creating date_time columns.
  df['Date/Time']= df['Start Date'].astype(str)+' '+df['Start Time'].astype(str)
  df['Date/Time'] = pd.to_datetime(df['Date/Time'],format=('%Y-%m-%d  %H:%M:%S'))
  df['Date/Time'] = [x.round('T') for x in df['Date/Time']]
  df["Nepal Date/Time"]=df['Date/Time']+timedelta(hours=9.75)
  df['Nepal Date/Time'] = [x.round('T') for x in df['Nepal Date/Time']]
  df = df.sort_values(by="Date/Time")

  df["Author_I"]=np.where(df['Type']=="Invited",df["presenter"]+" (Invited)",df["presenter"])
  df.loc[df['method']=="Virtual",'CDP/FYT/Virtual']="Virtual Presentation"
  df.loc[df['method']=="Poster",'CDP/FYT/Virtual']="Poster Presentation"
  df.loc[df['method'].isin(["CDP"]),'CDP/FYT/Virtual']="In-Person Presentation, CDP"
  df.loc[df['method'].isin(["Fayetteville"]),'CDP/FYT/Virtual']="In-Person Presentation, Fayetteville"

  return df


def get_filtered_data(df):
  """
  Generate filtered data based on division and method.
  """
  divs = df.division.unique().tolist()
  methods = df.method.unique().tolist()

  filtered_data = {}
  for division in divs:
    dfd = df.loc[df['division'].str.contains(division)].sort_values(by='Date/Time')
    filtered_data[division] = dfd
  
  for method in methods:
    dfm = df.loc[df['method'].str.contains(method)].sort_values(by='Date/Time')
    filtered_data[method] = dfm
  return filtered_data

def prep_list(df):
    df = df.sort_values(by="Date/Time")
    df['Date/Time']=df['Date/Time'].dt.strftime('%Y/%m/%d %I:%M %p')
    df['Nepal Date/Time']=df['Nepal Date/Time'].dt.strftime('%Y/%m/%d %I:%M %p')
    Names_L= df['Author_I'].values
    title_L = df['title'].values
    affiliations_L = df['affiliation'].values
    abstract_L = df['abstract'].values
    abstract_n_L=df['abstract_id'].values
    Date_L=df['Date/Time'].astype(str).to_list()
    Date_L_N=df['Nepal Date/Time'].astype(str).to_list()
    Location_L = df['CDP/FYT/Virtual'].values
    return Names_L,affiliations_L,title_L,abstract_L,abstract_n_L,Date_L,Date_L_N, Location_L

def create_html(df_name,html_name):

  prepend="""
  <html>
  <head>
  <style>
  #anpatable {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 100%;
  }

  #anpatable td, #anpatable th {
    border: 1px solid #ddd;
    padding: 8px;
  }

  #anpatable tr:{background-color: #fff;}
  #anpatable tr:nth-child(odd) {background-color: #D5DBdb;}
  #anpatable tr:nth-child(even) {background-color: #e5e8e8;}

  #anpatable tr:hover {background-color: #f8f9f9;}

  #anpatable th {
    padding-top: 8px;
    padding-bottom: 8px;
    text-align: left;
    background-color: #4CAF50;
    color: white;
    width: 83px;
  }

  </style>
  </head>
  <body>
  <h3> Please look below for detailed schedule. </h3></br>

    <table border="1" cellspacing="0" id="anpatable" cellpadding="2" width="680">
      <tbody>

  """


  postpend="""
  </tbody>
  </table>

  </body>
  </html>

  """
  Names_L,affiliations_L,title_L,abstract_L,abstract_n_L,Date_L,Date_L_N, Location_L =prep_list(df_name)
  fo = open(html_name, 'w',encoding="utf-8")
  fo.write(prepend)
  fo = open(html_name, 'a+',encoding="utf-8")
  for i in range(len(Names_L)):

      names_ = Names_L[i]
      title_ = title_L[i]
      affil_ = affiliations_L[i]
      abstract_ = abstract_L[i]
      abstract_n_=abstract_n_L[i]
      date_=Date_L[i]
      nepaldate_=Date_L_N[i]
      location_ = Location_L[i]

      html_code = '''

      <!-- ########################################################################## -->

      <tr>
                  <td width="158" colspan="2">
                      <p> Date/Time: <br>
                          ET:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{date_} </br>
                          Nepal: {nepaldate_}
                      </p>
                  </td>

      <td width="521" colspan="6">
          <p>Abstract Number: {abstract_n_} </p>

          <p>Presenting Author: {names_}
          </p>

          <p>Presenter's Affiliation: {affil_}
          </p>

          <p><strong>Title: </strong>{title_}</a></p>

          <p> <strong>Location: </strong>{location_}</p>

          <p class="show_hide">Show/Hide Abstract</p>
          <div class="show_on_click">{abstract_}</div>

      </td>
      </tr>

  '''.format(date_=date_,nepaldate_=nepaldate_,abstract_n_=abstract_n_,affil_=affil_,names_=names_,
              title_ = title_,abstract_=abstract_,location_=location_)
      fo.write(html_code)

  fo = open(html_name, 'a+',encoding="utf-8")
  fo.write(postpend)
  fo.close()

def main():
  filename = "2024/merged_v4.xlsx"
  sheet =  'July11230AM working good'

  # original data
  df_orig = load_data(filename, sheet_name=sheet)
  # processed data
  df      = process_data(df_orig)
  # filtered data
  filtered_ = get_filtered_data(df)

  print (f"\n\n\n filtered keys: {filtered_.keys()}")
  reqd_data = {'divisions': ['Data Science, Quantum Computing, Artificial Intelligence',
                            'Physics Education Research',
                            'Condensed Matter Physics and Material Science',
                            'Astronomy /Space Science /Cosmo Science/ Atmospheric Physics', 
                            'Atomic, Molecular, Optical and Plasma Physics',
                            'High Energy, Particle, and Nuclear Physics',
                            'Applied and Industrial Physics',
                            'Biological, Medical, Soft Matter and Chemical Physics'],
              'methods': ['CDP', 'Poster', 'Fayetteville']}
  mapping = {'Data Science, Quantum Computing, Artificial Intelligence':'data_science',
              'Physics Education Research':'per',
              'Condensed Matter Physics and Material Science':'cmp',
              'Astronomy /Space Science /Cosmo Science/ Atmospheric Physics':'astro', 
              'Atomic, Molecular, Optical and Plasma Physics':'amo',
              'High Energy, Particle, and Nuclear Physics':'high_energy',
              'Applied and Industrial Physics':'applied',
              'Biological, Medical, Soft Matter and Chemical Physics':'bio'}
  # save dfs for all required data
  all_dfs = {}

  for division in reqd_data['divisions']:
    all_dfs[division] = filtered_[division]

  for method in reqd_data['methods']:
    all_dfs[method] = filtered_[method]
  
  for table_name, df_ in all_dfs.items():
    name_ =  mapping.get(table_name, table_name.lower().replace(' ', ''))
    create_html(df_, html_name = join(os.getcwd(), '2024', f"html_code_2024_{name_}.html"))


if __name__ == '__main__':
  main()