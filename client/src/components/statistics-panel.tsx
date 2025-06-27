import React from 'react';
import { BarChart3, TrendingUp, Users, Eye } from 'lucide-react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Separator } from '@/components/ui/separator';
import { formatNumber } from '@/lib/utils';
import type { ProcessingStats } from '@shared/schema';

interface StatisticsPanelProps {
  statistics: ProcessingStats | null;
  isLoading?: boolean;
}

export function StatisticsPanel({ statistics, isLoading }: StatisticsPanelProps) {
  if (isLoading) {
    return (
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center space-x-2">
            <BarChart3 className="h-5 w-5" />
            <span>Статистика обработки</span>
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          {[...Array(7)].map((_, i) => (
            <div key={i} className="flex items-center justify-between animate-pulse">
              <div className="h-4 bg-gray-200 rounded w-32"></div>
              <div className="h-4 bg-gray-200 rounded w-16"></div>
            </div>
          ))}
        </CardContent>
      </Card>
    );
  }

  if (!statistics) {
    return (
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center space-x-2">
            <BarChart3 className="h-5 w-5" />
            <span>Статистика обработки</span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-sm text-gray-500 text-center py-8">
            Статистика будет доступна после обработки файла
          </p>
        </CardContent>
      </Card>
    );
  }

  const statsItems = [
    {
      label: 'Всего строк обработано',
      value: formatNumber(statistics.totalRows),
      icon: <BarChart3 className="h-4 w-4 text-blue-500" />
    },
    {
      label: 'Отзывы',
      value: formatNumber(statistics.reviewsCount),
      icon: <Users className="h-4 w-4 text-green-500" />
    },
    {
      label: 'Комментарии Топ-20',
      value: formatNumber(statistics.commentsCount),
      icon: <Users className="h-4 w-4 text-purple-500" />
    },
    {
      label: 'Активные обсуждения',
      value: formatNumber(statistics.activeDiscussionsCount),
      icon: <TrendingUp className="h-4 w-4 text-orange-500" />
    }
  ];

  const summaryItems = [
    {
      label: 'Суммарные просмотры',
      value: formatNumber(statistics.totalViews),
      icon: <Eye className="h-4 w-4 text-blue-600" />
    },
    {
      label: 'Доля вовлечений',
      value: `${statistics.engagementRate}%`,
      icon: <TrendingUp className="h-4 w-4 text-green-600" />
    },
    {
      label: 'Площадки с данными',
      value: `${statistics.platformsWithData}%`,
      icon: <BarChart3 className="h-4 w-4 text-purple-600" />
    }
  ];

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center space-x-2">
          <BarChart3 className="h-5 w-5" />
          <span>Статистика обработки</span>
        </CardTitle>
      </CardHeader>
      <CardContent className="space-y-4">
        {statsItems.map((item, index) => (
          <div key={index} className="stats-item">
            <div className="flex items-center space-x-2">
              {item.icon}
              <span className="text-sm text-gray-600">{item.label}</span>
            </div>
            <span className="text-sm font-medium text-gray-900">{item.value}</span>
          </div>
        ))}
        
        <Separator className="my-4" />
        
        {summaryItems.map((item, index) => (
          <div key={index} className="stats-item">
            <div className="flex items-center space-x-2">
              {item.icon}
              <span className="text-sm text-gray-600">{item.label}</span>
            </div>
            <span className="text-sm font-medium text-gray-900">{item.value}</span>
          </div>
        ))}
      </CardContent>
    </Card>
  );
}
