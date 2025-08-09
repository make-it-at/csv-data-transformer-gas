// performanceMonitor.gs
// Version 1.0
// Last updated: 2025-03-02

var PerformanceMonitor = {
  metrics: {},
  startTimes: {},
  
  startOperation(operationName) {
    this.startTimes[operationName] = new Date().getTime();
    this.initializeMetrics(operationName);
    LoggerManager.debug(`Operation started: ${operationName}`);
  },

  endOperation(operationName) {
    const endTime = new Date().getTime();
    const duration = endTime - (this.startTimes[operationName] || endTime);
    
    this.metrics[operationName].duration = duration;
    this.metrics[operationName].endTime = new Date().toISOString();
    
    this.logMetrics(operationName);
    this.saveMetrics(operationName);
  },

  initializeMetrics(operationName) {
    this.metrics[operationName] = {
      startTime: new Date().toISOString(),
      apiCalls: 0,
      dataProcessed: 0,
      errors: 0,
      warnings: 0,
      memoryUsage: 0
    };
  },

  incrementMetric(operationName, metricName, value = 1) {
    if (!this.metrics[operationName]) {
      this.initializeMetrics(operationName);
    }
    this.metrics[operationName][metricName] += value;
  },

  recordAPICall(operationName, endpoint, duration) {
    this.incrementMetric(operationName, 'apiCalls');
    this.metrics[operationName].lastApiCall = {
      endpoint,
      duration,
      timestamp: new Date().toISOString()
    };
  },

  recordDataProcessed(operationName, bytes) {
    this.incrementMetric(operationName, 'dataProcessed', bytes);
  },

  recordError(operationName, error) {
    this.incrementMetric(operationName, 'errors');
    if (!this.metrics[operationName].errorDetails) {
      this.metrics[operationName].errorDetails = [];
    }
    this.metrics[operationName].errorDetails.push({
      message: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString()
    });
  },

  recordWarning(operationName, message) {
    this.incrementMetric(operationName, 'warnings');
    if (!this.metrics[operationName].warningDetails) {
      this.metrics[operationName].warningDetails = [];
    }
    this.metrics[operationName].warningDetails.push({
      message,
      timestamp: new Date().toISOString()
    });
  },

  getMetrics(operationName) {
    return this.metrics[operationName] || null;
  },

  getAllMetrics() {
    return this.metrics;
  },

  logMetrics(operationName) {
    const metrics = this.metrics[operationName];
    LoggerManager.info(`Performance metrics for ${operationName}`, {
      duration: `${metrics.duration}ms`,
      apiCalls: metrics.apiCalls,
      dataProcessed: `${(metrics.dataProcessed / 1024).toFixed(2)}KB`,
      errors: metrics.errors,
      warnings: metrics.warnings
    });
  },

  saveMetrics(operationName) {
    const metrics = this.metrics[operationName];
    const sheet = this.getMetricsSheet();
    
    sheet.appendRow([
      new Date(),
      operationName,
      metrics.duration,
      metrics.apiCalls,
      metrics.dataProcessed,
      metrics.errors,
      metrics.warnings,
      JSON.stringify(metrics)
    ]);
  },

  getMetricsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('PerformanceMetrics');
    
    if (!sheet) {
      sheet = ss.insertSheet('PerformanceMetrics');
      sheet.getRange('A1:H1').setValues([[
        'Timestamp',
        'Operation',
        'Duration (ms)',
        'API Calls',
        'Data Processed (bytes)',
        'Errors',
        'Warnings',
        'Raw Metrics'
      ]]);
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  },

  generateReport(operationName) {
    const metrics = this.metrics[operationName];
    if (!metrics) return null;

    return {
      summary: {
        operation: operationName,
        duration: metrics.duration,
        apiCalls: metrics.apiCalls,
        successRate: ((metrics.apiCalls - metrics.errors) / metrics.apiCalls * 100).toFixed(2) + '%',
        dataProcessed: (metrics.dataProcessed / 1024).toFixed(2) + 'KB'
      },
      details: {
        startTime: metrics.startTime,
        endTime: metrics.endTime,
        errors: metrics.errorDetails || [],
        warnings: metrics.warningDetails || []
      },
      recommendations: this.generateRecommendations(metrics)
    };
  },

  generateRecommendations(metrics) {
    const recommendations = [];

    if (metrics.duration > 300000) { // 5 minutes
      recommendations.push('Consider implementing batch processing to reduce execution time');
    }

    if (metrics.errors > 0) {
      recommendations.push('Implement retry logic for failed operations');
    }

    if (metrics.apiCalls > 100) {
      recommendations.push('Consider caching frequently accessed data');
    }

    return recommendations;
  },

  reset() {
    this.metrics = {};
    this.startTimes = {};
  },

  startTimer(operationName) {
    return {
      name: operationName,
      startTime: new Date().getTime(),
      checkpoints: []
    };
  },

  checkpoint(timer, checkpointName) {
    if (!timer) return;
    
    const currentTime = new Date().getTime();
    const elapsedSinceStart = currentTime - timer.startTime;
    const lastCheckpoint = timer.checkpoints.length > 0 ? 
      timer.checkpoints[timer.checkpoints.length - 1] : 
      { time: timer.startTime };
    const elapsedSinceLastCheckpoint = currentTime - lastCheckpoint.time;
    
    timer.checkpoints.push({
      name: checkpointName,
      time: currentTime,
      elapsedSinceStart,
      elapsedSinceLastCheckpoint
    });
  },

  endTimer(timer, logResults = true) {
    if (!timer) return null;
    
    const endTime = new Date().getTime();
    const totalTime = endTime - timer.startTime;
    
    const result = {
      name: timer.name,
      totalTime,
      checkpoints: timer.checkpoints,
      startTime: timer.startTime,
      endTime
    };
    
    if (logResults) {
      LoggerManager.info(`パフォーマンス [${timer.name}]: 合計時間=${totalTime}ms`);
      
      if (timer.checkpoints.length > 0) {
        timer.checkpoints.forEach(cp => {
          LoggerManager.debug(`  - ${cp.name}: ${cp.elapsedSinceLastCheckpoint}ms (累計: ${cp.elapsedSinceStart}ms)`);
        });
      }
    }
    
    return result;
  },

  analyzeBatchPerformance(batchResults) {
    if (!batchResults || batchResults.length === 0) {
      return null;
    }
    
    const totalBatches = batchResults.length;
    const totalTime = batchResults.reduce((sum, batch) => sum + batch.totalTime, 0);
    const avgBatchTime = totalTime / totalBatches;
    
    // 最速と最遅のバッチを特定
    const fastestBatch = batchResults.reduce((fastest, current) => 
      current.totalTime < fastest.totalTime ? current : fastest, batchResults[0]);
    
    const slowestBatch = batchResults.reduce((slowest, current) => 
      current.totalTime > slowest.totalTime ? current : slowest, batchResults[0]);
    
    const analysis = {
      totalBatches,
      totalTime,
      avgBatchTime,
      fastestBatch: {
        index: batchResults.indexOf(fastestBatch),
        time: fastestBatch.totalTime
      },
      slowestBatch: {
        index: batchResults.indexOf(slowestBatch),
        time: slowestBatch.totalTime
      }
    };
    
    LoggerManager.info(`バッチ処理分析: 合計=${totalBatches}バッチ, 平均時間=${avgBatchTime.toFixed(2)}ms`);
    LoggerManager.info(`  - 最速: バッチ#${analysis.fastestBatch.index + 1} (${analysis.fastestBatch.time}ms)`);
    LoggerManager.info(`  - 最遅: バッチ#${analysis.slowestBatch.index + 1} (${analysis.slowestBatch.time}ms)`);
    
    return analysis;
  }
};

// Usage functions
function recordPerformanceMetrics() {
  return PerformanceMonitor;
}

function getPerformanceReport(operationName) {
  return PerformanceMonitor.generateReport(operationName);
}