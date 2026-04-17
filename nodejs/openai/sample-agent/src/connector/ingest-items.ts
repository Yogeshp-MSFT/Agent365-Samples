// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Fetches recent NASA APOD entries and ingests them as external items
 * into the Microsoft Graph external connection created by create-connection.ts.
 *
 * Run:  npm run connector:ingest
 */

import { configDotenv } from 'dotenv';
configDotenv();

import { createGraphClient } from '../graph/graph-client';

const CONNECTION_ID = 'nasaApodConnector';
const NASA_API_KEY = process.env.NASA_API_KEY || 'DEMO_KEY';

interface ApodEntry {
  date: string;
  title: string;
  explanation: string;
  url: string;
  hdurl?: string;
  media_type: string;
  copyright?: string;
}

/**
 * Fetches the last `count` APOD entries from the NASA API with retry logic.
 */
async function fetchApodEntries(count: number): Promise<ApodEntry[]> {
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const res = await fetch(
      `https://api.nasa.gov/planetary/apod?api_key=${encodeURIComponent(NASA_API_KEY)}&count=${count}`
    );

    if (res.ok) {
      return res.json() as Promise<ApodEntry[]>;
    }

    if (res.status >= 500 && attempt < maxRetries) {
      const wait = attempt * 5;
      console.warn(`  NASA API returned ${res.status}. Retrying in ${wait}s… (attempt ${attempt}/${maxRetries})`);
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }

    throw new Error(`NASA API error: ${res.status} ${res.statusText}`);
  }

  throw new Error('NASA API: max retries exceeded');
}

/**
 * Returns hardcoded sample APOD entries as a fallback when the NASA API is unavailable.
 */
function getSampleApodEntries(): ApodEntry[] {
  return [
    {
      date: '2024-01-15',
      title: 'The Orion Nebula in Infrared',
      explanation: 'The Orion Nebula is one of the most observed and photographed objects in the night sky. Located about 1,344 light-years from Earth, it is the closest region of massive star formation to our planet.',
      url: 'https://apod.nasa.gov/apod/image/2401/OrionNebula_infrared.jpg',
      media_type: 'image',
      copyright: 'NASA/ESA/Hubble',
    },
    {
      date: '2024-02-20',
      title: 'Mars: Perseverance Rover Panorama',
      explanation: 'NASA\'s Perseverance rover captured this panoramic view of Jezero Crater on Mars, revealing ancient river delta deposits that once carried water and sediment into a lake billions of years ago.',
      url: 'https://apod.nasa.gov/apod/image/2402/MarsPanorama_perseverance.jpg',
      media_type: 'image',
    },
    {
      date: '2024-03-10',
      title: 'Pillars of Creation in Eagle Nebula',
      explanation: 'The Pillars of Creation are columns of cool interstellar gas and dust in the Eagle Nebula (M16), about 6,500 light-years from Earth. They were made famous by a Hubble Space Telescope image taken in 1995.',
      url: 'https://apod.nasa.gov/apod/image/2403/PillarsOfCreation_jwst.jpg',
      media_type: 'image',
      copyright: 'NASA/ESA/CSA/STScI',
    },
    {
      date: '2024-04-08',
      title: 'Total Solar Eclipse 2024',
      explanation: 'A total solar eclipse occurred on April 8, 2024, visible across North America. The Moon passed between the Earth and the Sun, briefly turning day into night and revealing the Sun\'s corona.',
      url: 'https://apod.nasa.gov/apod/image/2404/Eclipse2024_total.jpg',
      media_type: 'image',
    },
    {
      date: '2024-05-12',
      title: 'The Andromeda Galaxy in Ultraviolet',
      explanation: 'The Andromeda Galaxy (M31) is the nearest large galaxy to the Milky Way, located about 2.5 million light-years away. In ultraviolet light, its young hot stars and star-forming regions are highlighted.',
      url: 'https://apod.nasa.gov/apod/image/2405/AndromedaUV_galex.jpg',
      media_type: 'image',
      copyright: 'NASA/JPL-Caltech',
    },
    {
      date: '2024-06-21',
      title: 'Saturn\'s Rings in Detail',
      explanation: 'Saturn\'s rings are made of billions of particles of ice and rock, ranging in size from grains of sand to houses. Cassini spacecraft provided unprecedented close-up views during its Grand Finale orbits.',
      url: 'https://apod.nasa.gov/apod/image/2406/SaturnRings_cassini.jpg',
      media_type: 'image',
      copyright: 'NASA/JPL-Caltech/SSI',
    },
    {
      date: '2024-07-04',
      title: 'Crab Nebula Supernova Remnant',
      explanation: 'The Crab Nebula (M1) is the remnant of a supernova explosion recorded by Chinese astronomers in 1054 AD. At its center lies a rapidly spinning neutron star, or pulsar, that emits beams of radiation.',
      url: 'https://apod.nasa.gov/apod/image/2407/CrabNebula_jwst.jpg',
      media_type: 'image',
      copyright: 'NASA/ESA/CSA/STScI',
    },
    {
      date: '2024-08-15',
      title: 'Jupiter\'s Great Red Spot Close-Up',
      explanation: 'Jupiter\'s Great Red Spot is a persistent anticyclonic storm larger than Earth that has been observed for at least 350 years. Juno spacecraft revealed complex cloud structures within the massive storm.',
      url: 'https://apod.nasa.gov/apod/image/2408/JupiterRedSpot_juno.jpg',
      media_type: 'image',
      copyright: 'NASA/JPL-Caltech/SwRI/MSSS',
    },
    {
      date: '2024-09-22',
      title: 'Deep Field: Galaxies Across Time',
      explanation: 'The James Webb Space Telescope captured thousands of galaxies in this deep field image, some as they appeared over 13 billion years ago, offering a glimpse into the early universe.',
      url: 'https://apod.nasa.gov/apod/image/2409/DeepField_jwst.jpg',
      media_type: 'image',
      copyright: 'NASA/ESA/CSA/STScI',
    },
    {
      date: '2024-10-31',
      title: 'The Witch Head Nebula',
      explanation: 'The Witch Head Nebula (IC 2118) is a faint reflection nebula near the star Rigel in the constellation Orion. Its shape resembles a witch\'s profile, making it a popular target around Halloween.',
      url: 'https://apod.nasa.gov/apod/image/2410/WitchHead_nebula.jpg',
      media_type: 'image',
    },
  ];
}

/**
 * Pushes a single APOD entry into the Graph external connection.
 */
async function ingestItem(
  client: ReturnType<typeof createGraphClient>,
  entry: ApodEntry
): Promise<void> {
  const itemId = `apod-${entry.date}`;

  await client
    .api(`/external/connections/${CONNECTION_ID}/items/${itemId}`)
    .put({
      acl: [
        {
          type: 'everyone',
          value: 'everyone',
          accessType: 'grant',
        },
      ],
      properties: {
        title: entry.title,
        explanation: entry.explanation,
        date: entry.date,
        mediaType: entry.media_type,
        imageUrl: entry.hdurl || entry.url,
        copyright: entry.copyright || 'NASA / Public Domain',
      },
      content: {
        type: 'text',
        value: `${entry.title}\n\n${entry.explanation}`,
      },
    });
}

async function main(): Promise<void> {
  const count = Number(process.env.APOD_INGEST_COUNT) || 20;
  console.log(`Fetching ${count} random APOD entries from NASA…`);

  let entries: ApodEntry[];
  try {
    entries = await fetchApodEntries(count);
  } catch (err: any) {
    console.warn(`NASA API unavailable (${err.message}). Using built-in sample data instead.`);
    entries = getSampleApodEntries();
  }

  console.log(`Got ${entries.length} entries. Ingesting into Graph…`);

  const client = createGraphClient();

  for (const entry of entries) {
    try {
      await ingestItem(client, entry);
      console.log(`  ✓ ${entry.date} — ${entry.title}`);
    } catch (err: any) {
      console.error(`  ✗ ${entry.date} — ${err.message ?? err}`);
    }
  }

  console.log('Ingestion complete.');
}

main().catch((err) => {
  console.error('Ingestion failed:', err);
  process.exit(1);
});
