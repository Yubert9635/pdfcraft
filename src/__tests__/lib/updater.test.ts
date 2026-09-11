import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import {
  parseVersionParts,
  compareVersions,
  categorizeAsset,
  matchPlatformAssets,
  getUpdateSettings,
  saveUpdateSettings,
  shouldCheckUpdate,
  ignoreVersion,
  isVersionIgnored,
  checkUpdate,
} from '@/lib/updater';
import { ReleaseAsset } from '@/types/updater';

describe('Updater Library', () => {
  beforeEach(() => {
    localStorage.clear();
    vi.restoreAllMocks();
  });

  afterEach(() => {
    localStorage.clear();
    vi.restoreAllMocks();
  });

  describe('parseVersionParts', () => {
    it('parses standard semver strings', () => {
      expect(parseVersionParts('0.1.0').numbers).toEqual([0, 1, 0]);
      expect(parseVersionParts('v1.2.3').numbers).toEqual([1, 2, 3]);
    });

    it('parses date-based release tags with commit hashes', () => {
      expect(parseVersionParts('v2026.09.10-a1b2c3d').numbers).toEqual([2026, 9, 10]);
      expect(parseVersionParts('2026.9.1').numbers).toEqual([2026, 9, 1]);
    });

    it('handles non-numeric segments gracefully', () => {
      expect(parseVersionParts('abc.def').numbers).toEqual([0, 0]);
    });
  });

  describe('compareVersions', () => {
    it('returns true when latest semver is newer than current', () => {
      expect(compareVersions('0.1.0', '0.2.0')).toBe(true);
      expect(compareVersions('0.1.0', '1.0.0')).toBe(true);
      expect(compareVersions('1.2.3', '1.2.4')).toBe(true);
    });

    it('returns true when latest date tag is newer than current date tag', () => {
      expect(compareVersions('v2026.09.08-abc', 'v2026.09.10-def')).toBe(true);
      expect(compareVersions('2026.08.30', '2026.09.01')).toBe(true);
    });

    it('returns true when comparing initial 0.1.0 against a modern date release', () => {
      expect(compareVersions('0.1.0', 'v2026.09.10-abc')).toBe(true);
    });

    it('returns false when versions are identical', () => {
      expect(compareVersions('0.1.0', '0.1.0')).toBe(false);
      expect(compareVersions('v2026.09.10-abc', 'v2026.09.10-abc')).toBe(false);
    });

    it('returns false when current is newer than latest', () => {
      expect(compareVersions('0.2.0', '0.1.0')).toBe(false);
      expect(compareVersions('v2026.09.15', 'v2026.09.10')).toBe(false);
    });

    it('handles empty or missing versions safely', () => {
      expect(compareVersions('0.1.0', '')).toBe(false);
      expect(compareVersions('', 'v1.0.0')).toBe(true);
    });
  });

  describe('categorizeAsset', () => {
    it('identifies Windows portable zip', () => {
      expect(categorizeAsset('PDFCraft-Windows-x64-Portable.zip')).toBe('windows-portable');
      expect(categorizeAsset('pdfcraft-portable.zip')).toBe('windows-portable');
    });

    it('identifies Windows installer', () => {
      expect(categorizeAsset('PDFCraft_0.1.0_x64-setup.exe')).toBe('windows-installer');
      expect(categorizeAsset('PDFCraft_0.1.0_x64.msi')).toBe('windows-installer');
    });

    it('identifies macOS dmg', () => {
      expect(categorizeAsset('PDFCraft_0.1.0_x64.dmg')).toBe('macos-dmg');
    });

    it('identifies Linux packages', () => {
      expect(categorizeAsset('pdfcraft_0.1.0_amd64.AppImage')).toBe('linux-appimage');
      expect(categorizeAsset('pdfcraft_0.1.0_amd64.deb')).toBe('linux-deb');
    });

    it('identifies source or web export packages', () => {
      expect(categorizeAsset('pdfcraft-v2026.09.10.zip')).toBe('source');
      expect(categorizeAsset('source_code.tar.gz')).toBe('source');
    });

    it('returns other for unrecognized files', () => {
      expect(categorizeAsset('checksums.txt')).toBe('other');
    });
  });

  describe('matchPlatformAssets', () => {
    const mockAssets: ReleaseAsset[] = [
      {
        name: 'PDFCraft-Windows-x64-Portable.zip',
        downloadUrl: 'https://example.com/portable.zip',
        size: 50000000,
        platformType: 'windows-portable',
        browserDownloadUrl: 'https://example.com/portable.zip',
      },
      {
        name: 'PDFCraft_0.1.0_x64-setup.exe',
        downloadUrl: 'https://example.com/setup.exe',
        size: 55000000,
        platformType: 'windows-installer',
        browserDownloadUrl: 'https://example.com/setup.exe',
      },
      {
        name: 'PDFCraft_0.1.0_x64.dmg',
        downloadUrl: 'https://example.com/mac.dmg',
        size: 60000000,
        platformType: 'macos-dmg',
        browserDownloadUrl: 'https://example.com/mac.dmg',
      },
      {
        name: 'pdfcraft_0.1.0_amd64.AppImage',
        downloadUrl: 'https://example.com/linux.AppImage',
        size: 70000000,
        platformType: 'linux-appimage',
        browserDownloadUrl: 'https://example.com/linux.AppImage',
      },
      {
        name: 'pdfcraft_0.1.0_amd64.deb',
        downloadUrl: 'https://example.com/linux.deb',
        size: 65000000,
        platformType: 'linux-deb',
        browserDownloadUrl: 'https://example.com/linux.deb',
      },
    ];

    it('matches Windows platform with portable as primary', () => {
      const match = matchPlatformAssets(mockAssets, 'windows');
      expect(match.primary?.platformType).toBe('windows-portable');
      expect(match.secondary?.platformType).toBe('windows-installer');
      expect(match.all.length).toBe(5);
    });

    it('matches macOS platform with dmg as primary', () => {
      const match = matchPlatformAssets(mockAssets, 'macos');
      expect(match.primary?.platformType).toBe('macos-dmg');
    });

    it('matches Linux platform with AppImage as primary and deb as secondary', () => {
      const match = matchPlatformAssets(mockAssets, 'linux');
      expect(match.primary?.platformType).toBe('linux-appimage');
      expect(match.secondary?.platformType).toBe('linux-deb');
    });
  });

  describe('settings and ignored versions', () => {
    it('returns default settings when none are saved', () => {
      const settings = getUpdateSettings();
      expect(settings.autoCheck).toBe(true);
      expect(settings.checkFrequencyHours).toBe(24);
      expect(settings.ignoredVersions).toEqual([]);
    });

    it('saves and reads modified settings', () => {
      saveUpdateSettings({
        autoCheck: false,
        checkFrequencyHours: 48,
        lastCheckedTimestamp: 123456789,
        ignoredVersions: ['v2026.09.01'],
      });

      const updated = getUpdateSettings();
      expect(updated.autoCheck).toBe(false);
      expect(updated.checkFrequencyHours).toBe(48);
      expect(updated.lastCheckedTimestamp).toBe(123456789);
      expect(updated.ignoredVersions).toEqual(['v2026.09.01']);
    });

    it('adds and checks ignored versions', () => {
      expect(isVersionIgnored('v2026.09.10')).toBe(false);
      ignoreVersion('v2026.09.10');
      expect(isVersionIgnored('v2026.09.10')).toBe(true);
    });

    it('checks shouldCheckUpdate frequency correctly', () => {
      // Never checked -> true
      expect(shouldCheckUpdate()).toBe(true);

      // Just checked -> false
      const now = Date.now();
      saveUpdateSettings({
        autoCheck: true,
        checkFrequencyHours: 24,
        lastCheckedTimestamp: now,
        ignoredVersions: [],
      });
      expect(shouldCheckUpdate()).toBe(false);

      // Checked 25 hours ago -> true
      saveUpdateSettings({
        autoCheck: true,
        checkFrequencyHours: 24,
        lastCheckedTimestamp: now - 25 * 3600 * 1000,
        ignoredVersions: [],
      });
      expect(shouldCheckUpdate()).toBe(true);

      // Auto check disabled -> false
      saveUpdateSettings({
        autoCheck: false,
        checkFrequencyHours: 24,
        lastCheckedTimestamp: 0,
        ignoredVersions: [],
      });
      expect(shouldCheckUpdate()).toBe(false);
    });
  });

  describe('checkUpdate with mocked GitHub API', () => {
    it('fetches release and discovers new update successfully', async () => {
      const mockGitHubResponse = {
        tag_name: 'v2026.09.10-abc',
        name: 'PDFCraft v2026.09.10',
        published_at: '2026-09-10T12:00:00Z',
        html_url: 'https://github.com/PDFCraftTool/pdfcraft/releases/tag/v2026.09.10-abc',
        body: '## Release notes\n- Added in-app auto updater',
        assets: [
          {
            name: 'PDFCraft-Windows-x64-Portable.zip',
            browser_download_url: 'https://github.com/download/portable.zip',
            size: 10485760,
          },
        ],
      };

      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => mockGitHubResponse,
      });

      const result = await checkUpdate(true);

      expect(result.hasUpdate).toBe(true);
      expect(result.latestVersion).toBe('v2026.09.10-abc');
      expect(result.release?.name).toBe('PDFCraft v2026.09.10');
      expect(result.release?.body).toContain('Added in-app auto updater');
      expect(result.matchedAssets.primary?.name).toBe('PDFCraft-Windows-x64-Portable.zip');
    });

    it('handles network failure gracefully without throwing', async () => {
      global.fetch = vi.fn().mockRejectedValue(new Error('Network offline'));

      const result = await checkUpdate(true);

      expect(result.hasUpdate).toBe(false);
      expect(result.error).toContain('Network offline');
    });
  });
});
