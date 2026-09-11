'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { useTranslations } from 'next-intl';
import { ArrowUpCircle } from 'lucide-react';
import { checkUpdate, isVersionIgnored, shouldCheckUpdate } from '@/lib/updater';
import { isTauri } from '@/lib/tauri-bridge';
import { UpdateCheckResult } from '@/types/updater';
import { UpdateModal } from './UpdateModal';

export interface UpdateCheckButtonProps {
  className?: string;
}

export const UpdateCheckButton: React.FC<UpdateCheckButtonProps> = ({ className = '' }) => {
  const t = useTranslations('common');
  const [isOpen, setIsOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<UpdateCheckResult | null>(null);
  const [hasNotification, setHasNotification] = useState(false);

  const checkBtnLabel = (() => {
    try {
      const msg = t('updater.checkBtn' as any);
      if (msg && !msg.startsWith('updater.')) return msg;
    } catch {
      // fallback
    }
    return 'Check for updates';
  })();

  // Background auto-check on mount
  useEffect(() => {
    let isMounted = true;

    const performBackgroundCheck = async () => {
      // In Tauri or if auto-check interval has elapsed
      if (shouldCheckUpdate() || isTauri()) {
        try {
          const res = await checkUpdate(false);
          if (isMounted && res.hasUpdate && !isVersionIgnored(res.latestVersion)) {
            setResult(res);
            setHasNotification(true);
          }
        } catch {
          // Silent failure in background
        }
      }
    };

    // Small delay on initial startup to avoid contending with page hydration
    const timer = setTimeout(() => {
      performBackgroundCheck();
    }, 2500);

    return () => {
      isMounted = false;
      clearTimeout(timer);
    };
  }, []);

  const handleManualCheck = useCallback(async () => {
    setIsOpen(true);
    setLoading(true);
    try {
      const res = await checkUpdate(true);
      setResult(res);
      if (res.hasUpdate && !isVersionIgnored(res.latestVersion)) {
        setHasNotification(true);
      } else {
        setHasNotification(false);
      }
    } finally {
      setLoading(false);
    }
  }, []);

  const handleClose = () => {
    setIsOpen(false);
    setHasNotification(false);
  };

  return (
    <>
      <button
        type="button"
        onClick={handleManualCheck}
        className={`relative flex items-center justify-center h-9 w-9 rounded-lg text-[hsl(var(--color-muted-foreground))] hover:text-[hsl(var(--color-foreground))] hover:bg-[hsl(var(--color-muted))/0.5] transition-all ${className}`}
        aria-label={checkBtnLabel}
        title={checkBtnLabel}
      >
        <ArrowUpCircle className="h-5 w-5" aria-hidden="true" />

        {/* Pulse indicator for available updates */}
        {hasNotification && (
          <span className="absolute top-1.5 right-1.5 flex h-2.5 w-2.5">
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75" />
            <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500 shadow-sm" />
          </span>
        )}
      </button>

      <UpdateModal
        isOpen={isOpen}
        onClose={handleClose}
        result={result}
        loading={loading}
        onRetry={handleManualCheck}
      />
    </>
  );
};
