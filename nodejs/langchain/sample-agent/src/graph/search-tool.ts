// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * LangChain-compatible tool that searches the NASA APOD Graph Connector
 * via the Microsoft Graph Search API (/search/query).
 */

import { tool } from '@langchain/core/tools';
import { z } from 'zod';
import { createGraphClient } from './graph-client';

const CONNECTION_ID = 'nasaApodConnector';

interface SearchHit {
  resource: {
    properties?: {
      title?: string;
      explanation?: string;
      date?: string;
      mediaType?: string;
      imageUrl?: string;
      copyright?: string;
    };
  };
}

export const searchNasaApod = tool(
  async ({ query }: { query: string }): Promise<string> => {
    const client = createGraphClient();

    const searchResponse = await client.api('/search/query').post({
      requests: [
        {
          entityTypes: ['externalItem'],
          contentSources: [`/external/connections/${CONNECTION_ID}`],
          query: { queryString: query },
          from: 0,
          size: 5,
        },
      ],
    });

    const hits: SearchHit[] =
      searchResponse?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

    if (hits.length === 0) {
      return `No NASA APOD results found for "${query}".`;
    }

    const results = hits.map((hit, i) => {
      const props = hit.resource.properties ?? {};
      return [
        `**${i + 1}. ${props.title ?? 'Untitled'}** (${props.date ?? 'unknown date'})`,
        props.explanation
          ? props.explanation.length > 300
            ? props.explanation.slice(0, 300) + '…'
            : props.explanation
          : '',
        props.imageUrl ? `Image: ${props.imageUrl}` : '',
        props.copyright ? `© ${props.copyright}` : '',
      ]
        .filter(Boolean)
        .join('\n');
    });

    return results.join('\n\n');
  },
  {
    name: 'search_nasa_apod',
    description:
      'Search NASA Astronomy Picture of the Day entries indexed in Microsoft 365 via Graph Connectors. ' +
      'Use this when the user asks about space, astronomy, nebulae, galaxies, planets, or NASA images.',
    schema: z.object({
      query: z
        .string()
        .describe('The search keywords, e.g. "supernova", "Mars rover", "Orion nebula"'),
    }),
  }
);
