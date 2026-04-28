"use strict";

const ORG = process.env.ADO_ORG;
const PROJECT = process.env.ADO_PROJECT;
const PAT = process.env.ADO_PAT;
const BASE_URL = `https://dev.azure.com/${ORG}/${PROJECT}/_apis`;
const AUTH = Buffer.from(`:${PAT}`).toString("base64");
const HEADERS = {
  Authorization: `Basic ${AUTH}`,
  "Content-Type": "application/json",
};

module.exports = { ORG, PROJECT, BASE_URL, HEADERS };
