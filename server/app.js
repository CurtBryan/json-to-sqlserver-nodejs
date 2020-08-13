// * authored by Curt Bryan and Isaac Skidmore, 8-2020

const querystring = require("querystring");
const axios = require("axios");
const sql = require("mssql");
require("dotenv").config();

const SharePointKeyGrabConfig = {
  url: `https://accounts.accesscontrol.windows.net/${process.env.REALM}/tokens/OAuth/2`,
  data: {
    grant_type: process.env.GRANT_TYPE,
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    resource: process.env.RESOURCE,
  },
  headers: {
    "Content-Type": "application/x-www-form-urlencoded",
  },
};

const sqlConfig = {
  user: process.env.SQLUser,
  password: process.env.SQLPassword,
  server: process.env.SQLConnectionString,
  database: process.env.SQLDatabase,
};

const getSharePointListData = async () => {
  try {
    const { url, data, headers } = SharePointKeyGrabConfig;
    const keyGrab = await axios.post(url, querystring.stringify(data), headers);
    const sharepointList = await axios.get(process.env.SP_URL, {
      headers: {
        Authorization: `Bearer ${keyGrab.data.access_token}`,
        Accept: "application/json;odata=verbose",
      },
    });
    const parsedData = sharepointList.data.d.results.map((item) => {
      return {
        id: item.Id,
        title: item.Title,
        policyOwner: item.PolicyOwner,
        SMEStatus: item.SMEStatus,
      };
    });
    return parsedData;
  } catch (err) {
    console.error(err);
  }
};

const sendDataToSQLServer = async (parsedData) => {
  await sql.connect(sqlConfig, (err) => {
    if (err) return err.toString();
    const request = new sql.Request();
    parsedData.forEach((file) => {
      request
        .query(
          `insert into Data (Id, Title, PolicyOwner, SMEStatus) values ('${file.id}', '${file.title}', '${file.policyOwner}', '${file.SMEStatus}')`
        )
        .then((resp) => {
          return "ok";
        })
        .catch((err) => {
          return err.toString();
        });
    });
  });
  return "ok";
};

const clearSQLTable = async () => {
  await sql.connect(sqlConfig, (err) => {
    if (err) {
      return err.toString();
    }
    const request = new sql.Request();
    request
      .query("delete from Data")
      .then((resp) => {
        return "ok";
      })
      .catch((err) => {
        return err.toString();
      });
  });
  return "ok";
};

const runSQLJob = async () => {
  try {
    const sharePointData = await getSharePointListData();
    if (!sharePointData) throw "no data in SharePoint";
    const deleteDataFromSQL = await clearSQLTable();
    if (deleteDataFromSQL !== "ok") throw "deletion of SQL Data failed";
    const upsertToSQL = await sendDataToSQLServer(sharePointData);
    if (upsertToSQL !== "ok") throw "upsertion to SQL failed";
    console.log("ok")
    return;
  } catch (err) {
    console.error(err);
    return;
  }
};

runSQLJob();

module.exports = {
  listGrabTest: async () => {
    const results = await getSharePointListData();
    return results;
  },
  sendToSQLTest: async (data) => {
    const results = await sendDataToSQLServer(data);
    return results;
  },
  deleteFromSQL: async () => {
    const results = await clearSQLTable();
    return results;
  },
};
