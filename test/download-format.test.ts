import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import {
  AttachmentTicketStore,
  TICKET_PARAM,
  isPlainGraphDownloadPath,
  isPlainGraphPath,
} from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';

// get-download-url with ?format=pdf|html on a drive item's /content: Graph converts the file
// on download, and the server mints its own single-use URL that fetches exactly that path.

describe('isPlainGraphDownloadPath', () => {
  it.each([
    '/drives/d1/items/i1/content?format=pdf',
    '/drives/d1/items/i1/content?format=html',
    '/me/drive/items/i1/content?format=pdf',
    '/me/drive/root:/Forms/a.doc:/content?format=pdf',
  ])('accepts %s', (target) => {
    expect(isPlainGraphDownloadPath(target)).toBe(true);
  });

  it.each([
    '/drives/d1/items/i1/content?format=docx',
    '/drives/d1/items/i1/content?format=pdf&x=1',
    '/drives/d1/items/i1/content?$select=id',
    '/drives/d1/items/i1?format=pdf',
    '/drives/d1/items/../i2/content?format=pdf',
    '/drives/d1/items/i1/content?format=pdf#frag',
    '/drives/d1/items/i1/content?format=PDF',
  ])('refuses %s', (target) => {
    expect(isPlainGraphDownloadPath(target)).toBe(false);
  });

  it('leaves the plain check unchanged: no query of any kind', () => {
    expect(isPlainGraphPath('/drives/d1/items/i1/content?format=pdf')).toBe(false);
    expect(isPlainGraphDownloadPath('/drives/d1/items/i1/content')).toBe(true);
  });
});

describe('get-download-url with format=', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
  let store: AttachmentTicketStore;
  const ctx = {
    graphClient: {
      graphRequest: () => {
        throw new Error('must not call Graph to mint');
      },
    } as never,
    authManager: {
      isOAuthModeEnabled: () => false,
      isMultiAccount: async () => false,
      getToken: async () => 'T',
      getTokenForAccount: async () => 'T',
    } as never,
    multiAccount: false,
    accountNames: [],
  };
  const parse = (r: unknown) =>
    JSON.parse((r as { content: Array<{ text: string }> }).content[0].text);

  afterEach(() => resetAttachmentMinting());

  describe('with --enable-attachment-urls', () => {
    beforeEach(() => {
      store = new AttachmentTicketStore(120);
      configureAttachmentMinting({
        store,
        config: { base: 'http://m365:3000', key: 'k', keyId: '1', ttlSeconds: 120 },
      });
    });

    it('mints a single-use URL whose ticket fetches /content?format=pdf', async () => {
      const target =
        '/drives/66791E00B5BA8ECE/items/66791E00B5BA8ECE!s1e84be4b93f14dd690e7aebb4d78b6cb/content?format=pdf';
      const result = await tool.execute({ target }, ctx);
      expect((result as { isError?: boolean }).isError).toBeFalsy();
      const url = new URL(parse(result).downloadUrl);
      expect(url.origin + url.pathname).toBe('http://m365:3000/attachment');
      const ticket = store.redeem(url.searchParams.get(TICKET_PARAM)!);
      expect(ticket).toMatchObject({ target });
    });

    it.each([
      ['/drives/d1/items/i1/content?format=docx', 'format= is accepted only'],
      ['/drives/d1/items/i1?format=pdf', 'format= is accepted only'],
      ['/me/messages/m1/attachments/a1/$value?format=pdf', 'format= is accepted only'],
      ['/drives/d1/items/i1/content?$select=id', 'must not include query parameters'],
    ])('refuses %s', async (target, message) => {
      const result = await tool.execute({ target }, ctx);
      expect((result as { isError?: boolean }).isError).toBe(true);
      expect(parse(result).error).toContain(message);
    });
  });

  it('without --enable-attachment-urls, says why and mints nothing', async () => {
    const result = await tool.execute({ target: '/drives/d1/items/i1/content?format=pdf' }, ctx);
    expect((result as { isError?: boolean }).isError).toBe(true);
    expect(parse(result).error).toContain('--enable-attachment-urls');
  });
});
