const { listGrabTest, sendToSQLTest, deleteFromSQL } = require("./app.js");

const testData = [
  {
    id: 1,
    title: "Policy Cycle",
    policyOwner: "Ethics, Compliance, and Data Policy",
    SMEStatus: "Not Sent",
  },
  {
    id: 3,
    title: "Statement of Work Policy Library Solution",
    policyOwner: "Ethics, Compliance, and Data Policy",
    SMEStatus: "Approved",
  },
];

test("Test data grab from SharePoint", async () => {
  const results = await listGrabTest();
  expect(results).toBeDefined();
});

test("Test data push to SQL", async () => {
  const results = await sendToSQLTest(testData);
  expect(results).toBe("ok");
});

test("Test data delete to SQL", async () => {
  const results = await deleteFromSQL(testData);
  expect(results).toBe("ok");
});
