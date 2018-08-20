"""A simple example of how to access the Google Analytics API."""

from apiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import xlsxwriter as xl
import numpy as np
import pandas as pd


def get_service(api_name, api_version, scopes, key_file_location):
    """Get a service that communicates to a Google API.

    Args:
        api_name: The name of the api to connect to.
        api_version: The api version to connect to.
        scopes: A list auth scopes to authorize for the application.
        key_file_location: The path to a valid service account JSON key file.

    Returns:
        A service that is connected to the specified API.
    """

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
            key_file_location, scopes=scopes)

    # Build the service object.
    service = build(api_name, api_version, credentials=credentials)

    return service


def get_first_profile_id(service):
    # Use the Analytics service object to get the first profile id.

    # Get a list of all Google Analytics accounts for this user
    accounts = service.management().accounts().list().execute()

    if accounts.get('items'):
        # Get the first Google Analytics account.
        account = accounts.get('items')[0].get('id')

        # Get a list of all the properties for the first account.
        properties = service.management().webproperties().list(
                accountId=account).execute()

        if properties.get('items'):
            # Get the first property id.
            property = properties.get('items')[0].get('id')

            # Get a list of all views (profiles) for the first property.
            profiles = service.management().profiles().list(
                    accountId=account,
                    webPropertyId=property).execute()

            if profiles.get('items'):
                # return the first view (profile) id.
                return profiles.get('items')[0].get('id')

    return None


def get_results2(service, profile_id):
    # Use the Analytics Service Object to query the Core Reporting API
    # for the number of sessions within the past seven days.
    locations = service.data().ga().get(
            ids='ga:' + profile_id,
            start_date='7daysAgo',
            end_date='today',
            dimensions='ga:country, ga:region, ga:city, ga:pagePath',
            metrics='ga:users, ga:sessions, ga:sessionDuration').execute()

    pviews = []
    for page in locations['rows']:
        if page[0] == 'United States' or page[0] == 'Canada' or page[0] == 'Mexico':
            pviews.append(page)
    return pviews

def create_df(data1):
    headers = ['Country', 'State', 'City', 'Pages Visited','Total Users','Total Sessions', 'Total Time On Page']
    df = pd.DataFrame(data1, columns = headers )
    df[['Total Time On Page']] = df[['Total Time On Page']].apply(pd.to_numeric)

    df60 = df[(df['Pages Visited'] != '/') & (df['Total Time On Page'] > 0)].sort_values('Total Time On Page', ascending = False).reset_index(drop = True)
    df60['Total Time On Page'] = df60['Total Time On Page']/60
    return df60

def write_to_excel(df):

    new_file = pd.ExcelWriter('WebsiteHits.xlsx')
    df.to_excel(new_file, index=False, sheet_name='Website Visits Last 7 Days')
    wb = new_file.book
    ws = new_file.sheets['Website Visits Last 7 Days']
    center = wb.add_format({'align': 'center'})
    bold = wb.add_format({'bold': True, 'align': 'center'})
    ws.set_zoom(90)

    wb.close()

def main():
    # Define the auth scopes to request.
    scope = 'https://www.googleapis.com/auth/analytics.readonly'
    key_file_location = 'client_secrets_fx.json'

    # Authenticate and construct service.
    service = get_service(
            api_name='analytics',
            api_version='v3',
            scopes=[scope],
            key_file_location=key_file_location)

    profile_id = get_first_profile_id(service)
    location_results2 = get_results2(service, profile_id)
    df = create_df(location_results2)
    #print(len(location_results))
    write_to_excel(df)


if __name__ == '__main__':
    main()
