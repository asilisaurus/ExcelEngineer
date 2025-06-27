import React from 'react';
import { cn } from '@/lib/utils';

interface ProgressBarProps {
  value: number; // 0-100
  className?: string;
  showLabel?: boolean;
  size?: 'sm' | 'md' | 'lg';
  variant?: 'default' | 'success' | 'warning' | 'error';
}

export function ProgressBar({ 
  value, 
  className, 
  showLabel = true, 
  size = 'md',
  variant = 'default'
}: ProgressBarProps) {
  const sizeClasses = {
    sm: 'h-1',
    md: 'h-2',
    lg: 'h-3'
  };

  const variantClasses = {
    default: 'bg-primary',
    success: 'bg-green-500',
    warning: 'bg-yellow-500',
    error: 'bg-red-500'
  };

  const clampedValue = Math.min(Math.max(value, 0), 100);

  return (
    <div className={cn("space-y-2", className)}>
      {showLabel && (
        <div className="flex items-center justify-between text-sm">
          <span className="text-gray-600">Прогресс обработки</span>
          <span className="font-medium text-gray-900">{Math.round(clampedValue)}%</span>
        </div>
      )}
      <div className={cn("w-full bg-gray-200 rounded-full", sizeClasses[size])}>
        <div
          className={cn(
            "rounded-full transition-all duration-500 ease-out",
            sizeClasses[size],
            variantClasses[variant]
          )}
          style={{ width: `${clampedValue}%` }}
        />
      </div>
    </div>
  );
}
