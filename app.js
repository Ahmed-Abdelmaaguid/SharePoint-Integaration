const express = require('express');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const axios = require('axios');
const AdmZip = require('adm-zip');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
require('dotenv').config();

const app = express();

app.use(cors());
app.use(express.json());

// SharePoint configuration
const {
  CLIENT_ID: clientId,
  CLIENT_SECRET: clientSecret,
  TENANT_ID: tenantId,
  SITE_ID: siteId,
  DRIVE_ID: driveId,
  URL: url,
  USER: username,
  PASSWORD: password,
  ALIAS: alias,
  COMPANYID: companyid,
  APITOKEN: apitoken
} = process.env;

async function downloadAndExtract(url, params) {
  try {
    console.log('The params are', params);
    const response = await axios.post(url, null, {
      params,
      responseType: 'arraybuffer',
      headers: {
        'Accept': 'application/zip'
      }
    });

    if (response.status === 200) {
      if (response.data.length === 0) {
        console.log('Received an empty response from the server');
        return null;
      }
      console.log('Data received:', response.data);
      const zip = new AdmZip(response.data);
      return zip;
    } else {
      console.log(`Failed to retrieve the ZIP file. Status code: ${response.status}`);
      return null;
    }
  } catch (error) {
    console.error("Error downloading the ZIP file:", error.message);
    return null;
  }
}

function createZipBuffer(zip) {
  return zip.toBuffer();
}

async function uploadToSharePoint(filename, fileBuffer) {
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
  const accessToken = tokenResponse.token;

  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => accessToken
    }
  });

  try {
    const uploadSession = await client.api(`/sites/${siteId}/drives/${driveId}/root:/${filename}:/createUploadSession`)
      .post({
        item: {
          "@microsoft.graph.conflictBehavior": "replace"
        }
      });

    // console.log('Upload session details', uploadSession);
    const maxSliceSize = 320 * 1024; // 320 KB chunk size
    let start = 0;

    while (start < fileBuffer.length) {
      const end = Math.min(start + maxSliceSize, fileBuffer.length);
      const slice = fileBuffer.slice(start, end);

      await axios.put(uploadSession.uploadUrl, slice, {
        headers: {
          'Content-Length': slice.length,
          'Content-Range': `bytes ${start}-${end - 1}/${fileBuffer.length}`,
          'Authorization': `Bearer ${accessToken}`
        }
      });

      start = end;
    }

    console.log(`Upload of ${filename} completed.`);

    const fileMetadata = await client.api(`/sites/${siteId}/drives/${driveId}/root:/${filename}`).get();
    // console.log('File metadata:', fileMetadata);
    return fileMetadata;

  } catch (error) {
    console.error("Error uploading to SharePoint:", error.message);
    throw error;
  }
}

app.post('/api/upload', async (req, res) => {

  const referer = req.headers['referer'];
  const origin = req.headers['origin'];
  
  console.log('API called from Referer:', referer);
  console.log('API called from Origin:', origin);

  const { objectid, files } = req.body;
  const successfulUploads = [];
  const failedUploads = [];

  try {
    for (const file of files) {
      const fieldId = file.fieldid;
      const filename = file.filename;
      const fileExtension = path.extname(filename).toLowerCase();
      const params = {
        username,
        alias,
        companyid,
        password,
        objectid,
        fieldid: fieldId,
        filename,
        apitoken
      };

      let fileBuffer;

      try {
        if (fileExtension === '.pdf') {
          const response = await axios.post(url, null, { params, responseType: 'arraybuffer' });
          if (response.status === 200) {
            fileBuffer = Buffer.from(response.data);
          } else {
            console.log(`Failed to retrieve the file. Status code: ${response.status}`);
            failedUploads.push({ filename, error: `Failed to retrieve the file. Status code: ${response.status}` });
            continue;
          }
        } else if (fileExtension === '.docx' || fileExtension === '.xlsx') {
          const zip = await downloadAndExtract(url, params);
          if (zip) {
            fileBuffer = createZipBuffer(zip);
          } else {
            console.log(`Failed to process ${fileExtension} file.`);
            failedUploads.push({ filename, error: `Failed to process ${fileExtension} file.` });
            continue;
          }
        } else {
          console.log(`Unsupported file type: ${fileExtension}`);
          failedUploads.push({ filename, error: `Unsupported file type: ${fileExtension}` });
          continue;
        }

        await uploadToSharePoint(filename, fileBuffer);
        successfulUploads.push(filename);

      } catch (error) {
        console.error(`Error processing file ${filename}:`, error.message);
        failedUploads.push({ filename, error: error.message });
      }
    }

    res.json({ status: "success", successfulUploads, failedUploads });
  } catch (error) {
    res.status(500).json({ status: "failure", error: error.message });
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
