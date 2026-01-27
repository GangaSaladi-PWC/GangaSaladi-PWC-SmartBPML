// Test script to verify SAP Fiori Library API calls
// Run: node test-fiori-api.js

const testFioriAPI = async () => {
  const releaseId = 'S31PCE'; // SAP S/4HANA 2023
  const tcode = 'VA01'; // Sales Order transaction

  const fioriApiUrl = 'https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata';
  const url = `${fioriApiUrl}/AppListResult(sWhereClause='(1=1) and ("releaseId" = ''${releaseId}'') and ("TRANSACTION_MATCH" = ''${tcode}'')',INPLANGUAGE='None',sUUID='')/Results?$top=100&$select=fioriId,AppName,ApplicationType,TRANSACTION_MATCH,GTMAppDescription,UITechnology&$format=json`;

  console.log('Testing SAP Fiori Library API...\n');
  console.log('URL:', url);
  console.log('\nFetching...\n');

  try {
    const response = await fetch(url);

    if (!response.ok) {
      console.error('❌ API call failed with status:', response.status);
      return;
    }

    const data = await response.json();
    const results = data.d?.results || [];

    console.log(`✅ Total apps found: ${results.length}\n`);

    // Filter out SAP GUI
    const filteredApps = results.filter(app =>
      app.ApplicationType && app.ApplicationType !== 'SAP GUI'
    );

    console.log(`✅ After filtering SAP GUI: ${filteredApps.length}\n`);

    console.log('📋 Apps Details:\n');
    filteredApps.forEach((app, idx) => {
      console.log(`${idx + 1}. ${app.AppName}`);
      console.log(`   Fiori ID: ${app.fioriId}`);
      console.log(`   Type: ${app.ApplicationType}`);
      console.log(`   UI Tech: ${app.UITechnology}`);
      console.log('');
    });

    // Show how they would be concatenated
    console.log('📝 Comma-separated values (as they appear in Excel):\n');
    console.log('Fiori ID:', filteredApps.map(app => app.fioriId).join(', '));
    console.log('Fiori App Name:', filteredApps.map(app => app.AppName).join(', '));
    console.log('Application Type:', filteredApps.map(app => app.ApplicationType).join(', '));

  } catch (error) {
    console.error('❌ Error:', error.message);
  }
};

testFioriAPI();
