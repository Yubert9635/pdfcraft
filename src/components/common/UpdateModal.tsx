'use client';

import React, { useState } from 'react';
import { useTranslations } from 'next-intl';
import {
  Download,
  ExternalLink,
  CheckCircle2,
  AlertCircle,
  RefreshCw,
  Sparkles,
  Package,
  ChevronDown,
  ChevronUp,
} from 'lucide-react';
import { Modal } from '@/components/ui/Modal';
import { Button } from '@/components/ui/Button';
import { UpdateCheckResult, ReleaseAsset } from '@/types/updater';
import { ignoreVersion, GITHUB_REPO } from '@/lib/updater';

export interface UpdateModalProps {
  isOpen: boolean;
  onClose: () => void;
  result: UpdateCheckResult | null;
  loading: boolean;
  onRetry: () => void;
}

export const UpdateModal: React.FC<UpdateModalProps> = ({
  isOpen,
  onClose,
  result,
  loading,
  onRetry,
}) => {
  const t = useTranslations('common');
  const [showAllAssets, setShowAllAssets] = useState(false);

  const getMsg = (
    key: string,
    fallback: string,
    values?: Record<string, string | number>
  ): string => {
    try {
      const fullKey = `updater.${key}`;
      const msg = t(fullKey as any, values as any);
      if (msg && !msg.startsWith('updater.')) return msg;
    } catch {
      // fallback
    }
    if (values) {
      let res = fallback;
      for (const [k, v] of Object.entries(values)) {
        res = res.replace(`{${k}}`, String(v));
      }
      return res;
    }
    return fallback;
  };

  const formatFileSize = (bytes: number): string => {
    if (!bytes) return '';
    const mb = bytes / (1024 * 1024);
    return `(${mb.toFixed(1)} MB)`;
  };

  const getAssetLabel = (asset: ReleaseAsset): string => {
    switch (asset.platformType) {
      case 'windows-portable':
        return getMsg('downloadPortable', 'Windows Portable (ZIP)');
      case 'windows-installer':
        return getMsg('downloadInstaller', 'Windows Installer (.exe / .msi)');
      case 'macos-dmg':
        return getMsg('downloadMac', 'macOS (.dmg)');
      case 'linux-appimage':
        return getMsg('downloadAppImage', 'Linux AppImage (Portable)');
      case 'linux-deb':
        return getMsg('downloadDeb', 'Linux Debian (.deb)');
      default:
        return asset.name;
    }
  };

  const handleDownload = (asset: ReleaseAsset, useMirror = false) => {
    if (typeof window !== 'undefined') {
      const url = useMirror && asset.mirrorDownloadUrl ? asset.mirrorDownloadUrl : asset.browserDownloadUrl;
      window.open(url, '_blank', 'noopener,noreferrer');
    }
  };

  const handleSkipVersion = () => {
    if (result?.latestVersion) {
      ignoreVersion(result.latestVersion);
    }
    onClose();
  };

  return (
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      title={getMsg('title', 'Software Update')}
      size="lg"
    >
      <div className="space-y-5 py-2">
        {/* State: Loading */}
        {loading && (
          <div className="flex flex-col items-center justify-center py-10 space-y-4">
            <RefreshCw className="h-9 w-9 animate-spin text-[hsl(var(--color-primary))]" />
            <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
              {getMsg('checking', 'Checking for updates...')}
            </p>
          </div>
        )}

        {/* State: Error */}
        {!loading && result?.error && (
          <div className="flex flex-col items-center justify-center py-8 text-center space-y-4">
            <div className="flex h-12 w-12 items-center justify-center rounded-full bg-red-100 dark:bg-red-950 text-red-600 dark:text-red-400">
              <AlertCircle className="h-6 w-6" />
            </div>
            <div className="space-y-1">
              <h3 className="font-semibold text-[hsl(var(--color-foreground))]">
                {getMsg('errorTitle', 'Unable to check for updates')}
              </h3>
              <p className="text-xs text-[hsl(var(--color-muted-foreground))] max-w-sm">
                {result.error}
              </p>
            </div>
            <div className="flex flex-wrap items-center justify-center gap-2.5 pt-1">
              <Button variant="outline" size="sm" onClick={onRetry} className="gap-2">
                <RefreshCw className="h-4 w-4" />
                {getMsg('retry', 'Retry')}
              </Button>
              <Button
                variant="primary"
                size="sm"
                className="gap-1.5"
                onClick={() => {
                  if (typeof window !== 'undefined') {
                    window.open(`https://github.com/${GITHUB_REPO}/releases`, '_blank', 'noopener,noreferrer');
                  }
                }}
              >
                <ExternalLink className="h-3.5 w-3.5" />
                <span>{getMsg('viewOnGithub', 'View on GitHub')}</span>
              </Button>
            </div>
          </div>
        )}

        {/* State: Up to date */}
        {!loading && !result?.error && !result?.hasUpdate && (
          <div className="flex flex-col items-center justify-center py-8 text-center space-y-4">
            <div className="flex h-12 w-12 items-center justify-center rounded-full bg-emerald-100 dark:bg-emerald-950 text-emerald-600 dark:text-emerald-400">
              <CheckCircle2 className="h-6 w-6" />
            </div>
            <div className="space-y-1">
              <h3 className="text-base font-semibold text-[hsl(var(--color-foreground))]">
                {getMsg('latest', 'You are up to date!')}
              </h3>
              <p className="text-sm text-[hsl(var(--color-muted-foreground))]">
                {getMsg(
                  'latestDesc',
                  'PDFCraft is currently at the latest version ({version}).',
                  { version: result?.currentVersion || '' }
                )}
              </p>
            </div>
            <div className="pt-2">
              <Button variant="primary" size="sm" onClick={onClose}>
                {getMsg('close', 'Close')}
              </Button>
            </div>
          </div>
        )}

        {/* State: Update available */}
        {!loading && !result?.error && result?.hasUpdate && (
          <div className="space-y-5">
            {/* Version banner */}
            <div className="flex items-start gap-4 rounded-xl border border-[hsl(var(--color-primary))/0.25] bg-[hsl(var(--color-primary))/0.05] p-4">
              <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-[hsl(var(--color-primary))] text-white shadow-md">
                <Sparkles className="h-5 w-5" />
              </div>
              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2 flex-wrap">
                  <h3 className="font-bold text-[hsl(var(--color-foreground))] text-base">
                    {getMsg('available', 'New Version Available')}
                  </h3>
                  <span className="rounded-full bg-emerald-100 dark:bg-emerald-900/50 px-2.5 py-0.5 text-xs font-semibold text-emerald-700 dark:text-emerald-300">
                    {result.latestVersion}
                  </span>
                </div>
                <p className="text-xs text-[hsl(var(--color-muted-foreground))] mt-1">
                  {getMsg('currentVersion', 'Current')}:{' '}
                  <span className="font-mono">{result.currentVersion}</span>
                  {result.release?.publishedAt && (
                    <span className="ml-3">
                      {getMsg('released', 'Released')}:{' '}
                      {new Date(result.release.publishedAt).toLocaleDateString()}
                    </span>
                  )}
                </p>
              </div>
            </div>

            {/* Release Notes */}
            {result.release?.body && (
              <div className="space-y-2">
                <div className="text-xs font-semibold uppercase tracking-wider text-[hsl(var(--color-muted-foreground))]">
                  {getMsg('releaseNotes', 'Release Notes')}
                </div>
                <div className="max-h-48 overflow-y-auto rounded-lg border border-[hsl(var(--color-border))] bg-[hsl(var(--color-muted))/0.3] p-3 text-xs text-[hsl(var(--color-foreground))] whitespace-pre-wrap font-sans leading-relaxed">
                  {result.release.body}
                </div>
              </div>
            )}

            {/* Download Actions */}
            <div className="space-y-2.5 pt-1">
              {result.matchedAssets.primary && (
                <Button
                  variant="primary"
                  className="w-full justify-center gap-2 py-5 text-sm font-semibold shadow-md shadow-primary/20"
                  onClick={() => handleDownload(result.matchedAssets.primary!)}
                >
                  <Download className="h-4 w-4" />
                  <span>{getAssetLabel(result.matchedAssets.primary)}</span>
                  <span className="opacity-75 font-normal text-xs">
                    {formatFileSize(result.matchedAssets.primary.size)}
                  </span>
                </Button>
              )}

              {result.matchedAssets.secondary && (
                <Button
                  variant="outline"
                  className="w-full justify-center gap-2 py-4 text-xs"
                  onClick={() => handleDownload(result.matchedAssets.secondary!)}
                >
                  <Package className="h-3.5 w-3.5" />
                  <span>{getAssetLabel(result.matchedAssets.secondary)}</span>
                  <span className="opacity-75 font-mono text-[11px]">
                    {formatFileSize(result.matchedAssets.secondary.size)}
                  </span>
                </Button>
              )}

              {/* Fast mirror option */}
              {result.matchedAssets.primary?.mirrorDownloadUrl && (
                <div className="flex items-center justify-center pt-0.5">
                  <button
                    type="button"
                    onClick={() => handleDownload(result.matchedAssets.primary!, true)}
                    className="text-[11px] text-[hsl(var(--color-primary))] hover:underline inline-flex items-center gap-1 font-medium py-1 px-2 rounded hover:bg-[hsl(var(--color-primary))/0.08] transition-colors"
                  >
                    <span>⚡ 国内高速镜像通道下载</span>
                  </button>
                </div>
              )}

              {/* Toggle all assets */}
              {result.matchedAssets.all.length > 0 && (
                <div className="pt-1">
                  <button
                    type="button"
                    onClick={() => setShowAllAssets(!showAllAssets)}
                    className="flex items-center gap-1 text-xs text-[hsl(var(--color-muted-foreground))] hover:text-[hsl(var(--color-foreground))] transition-colors"
                  >
                    <span>{getMsg('allDownloads', 'View all platform downloads')}</span>
                    {showAllAssets ? (
                      <ChevronUp className="h-3.5 w-3.5" />
                    ) : (
                      <ChevronDown className="h-3.5 w-3.5" />
                    )}
                  </button>

                  {showAllAssets && (
                    <div className="mt-2 space-y-1.5 rounded-lg border border-[hsl(var(--color-border))] p-2 bg-[hsl(var(--color-background))]">
                      {result.matchedAssets.all.map((asset) => (
                        <div
                          key={asset.name}
                          className="flex items-center justify-between py-1 px-2 hover:bg-[hsl(var(--color-muted))/0.5] rounded transition-colors text-xs"
                        >
                          <span className="font-mono truncate max-w-[260px] text-[hsl(var(--color-foreground))]">
                            {asset.name}
                          </span>
                          <Button
                            variant="ghost"
                            size="sm"
                            className="h-7 px-2 gap-1 text-xs"
                            onClick={() => handleDownload(asset)}
                          >
                            <Download className="h-3 w-3" />
                            {formatFileSize(asset.size)}
                          </Button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Bottom auxiliary links */}
            <div className="flex items-center justify-between pt-3 border-t border-[hsl(var(--color-border))] text-xs">
              <a
                href={result.release?.htmlUrl || 'https://github.com/PDFCraftTool/pdfcraft/releases'}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1 text-[hsl(var(--color-primary))] hover:underline"
              >
                <span>{getMsg('viewOnGithub', 'View on GitHub')}</span>
                <ExternalLink className="h-3 w-3" />
              </a>

              <div className="flex items-center gap-2">
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={handleSkipVersion}
                  className="text-xs text-[hsl(var(--color-muted-foreground))] hover:text-[hsl(var(--color-foreground))]"
                >
                  {getMsg('skipVersion', 'Skip this version')}
                </Button>
                <Button variant="outline" size="sm" onClick={onClose} className="text-xs">
                  {getMsg('remindLater', 'Remind later')}
                </Button>
              </div>
            </div>
          </div>
        )}
      </div>
    </Modal>
  );
};
