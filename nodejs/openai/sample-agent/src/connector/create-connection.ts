// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Creates a Microsoft Graph external connection and registers the schema
 * for NASA APOD (Astronomy Picture of the Day) data.
 *
 * Run:  npm run connector:create
 */

import { configDotenv } from 'dotenv';
configDotenv();

import { createGraphClient } from '../graph/graph-client';

const CONNECTION_ID = 'nasaApodConnector';
const CONNECTION_NAME = 'NASA Astronomy Picture of the Day';
const CONNECTION_DESCRIPTION =
  'External data from NASA APOD – daily astronomy images with titles and explanations.';

async function createConnection(): Promise<void> {
  const client = createGraphClient();

  // 1. Create the external connection
  console.log(`Creating external connection "${CONNECTION_ID}"…`);
  try {
    await client.api('/external/connections').post({
      id: CONNECTION_ID,
      name: CONNECTION_NAME,
      description: CONNECTION_DESCRIPTION,
    });
    console.log('Connection created.');
  } catch (err: any) {
    if (err?.statusCode === 409) {
      console.log('Connection already exists — skipping creation.');
    } else {
      throw err;
    }
  }

  // 2. Register the schema
  console.log('Registering schema (this can take a few minutes)…');
  await client.api(`/external/connections/${CONNECTION_ID}/schema`).patch({
    baseType: 'microsoft.graph.externalItem',
    properties: [
      {
        name: 'title',
        type: 'String',
        isSearchable: true,
        isQueryable: true,
        isRetrievable: true,
        labels: ['title'],
      },
      {
        name: 'explanation',
        type: 'String',
        isSearchable: true,
        isRetrievable: true,
      },
      {
        name: 'date',
        type: 'String',
        isQueryable: true,
        isRetrievable: true,
      },
      {
        name: 'mediaType',
        type: 'String',
        isQueryable: true,
        isRetrievable: true,
      },
      {
        name: 'imageUrl',
        type: 'String',
        isRetrievable: true,
        labels: ['url'],
      },
      {
        name: 'copyright',
        type: 'String',
        isRetrievable: true,
      },
    ],
  });
  console.log('Schema registration requested. It may take 5-10 minutes to provision.');
  console.log('Run `npm run connector:ingest` once the schema is ready.');
}

createConnection().catch((err) => {
  console.error('Failed to create connection:', err);
  process.exit(1);
});
