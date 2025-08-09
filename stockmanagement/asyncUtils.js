// asyncUtils.gs
// Version 1.1
// Last updated: 2025-03-02
// Fixed for Google Apps Script environment

const AsyncUtils = {
  processingQueue: [],
  isProcessing: false,
  
  async processInBatches(items, batchSize, processor, options = {}) {
    const {
      maxConcurrent = 3,
      delayBetweenBatches = 1000
    } = options;

    const batches = this.createBatches(items, batchSize);
    const results = [];
    
    for (let i = 0; i < batches.length; i += maxConcurrent) {
      const batchPromises = batches
        .slice(i, i + maxConcurrent)
        .map(batch => processor(batch));
      
      const batchResults = await Promise.all(batchPromises);
      results.push(...batchResults);
      
      if (i + maxConcurrent < batches.length) {
        await this.delay(delayBetweenBatches);
      }
      
      if (this.shouldCancel()) break;
    }
    
    return results;
  },

  throttle(func, limit = 2) {
    let inThrottle;
    let lastResult;
    return async function(...args) {
      if (!inThrottle) {
        inThrottle = true;
        lastResult = await func.apply(this, args);
        await this.delay(1000 / limit);
        inThrottle = false;
      }
      return lastResult;
    }.bind(this);
  },

  createBatches(items, size) {
    return Array.from({ length: Math.ceil(items.length / size) }, (_, i) =>
      items.slice(i * size, (i + 1) * size)
    );
  },

  delay(ms) {
    Utilities.sleep(ms);
  },

  shouldCancel() {
    return PropertiesService.getScriptProperties()
      .getProperty('cancelFlag') === 'true';
  },

  async withRetry(func, options = {}) {
    const {
      maxAttempts = 3,
      baseDelay = 1000,
      maxDelay = 5000
    } = options;

    let lastError;
    
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        return await func();
      } catch (error) {
        lastError = error;
        if (attempt === maxAttempts) break;
        
        const delay = Math.min(
          baseDelay * Math.pow(2, attempt - 1),
          maxDelay
        );
        this.delay(delay);
      }
    }
    
    throw lastError;
  },
  
  async executeWithTimeout(func, timeoutMs = 300000) {
    const startTime = new Date().getTime();
    
    try {
      return await func();
    } catch (error) {
      if (new Date().getTime() - startTime >= timeoutMs) {
        throw new Error('Operation timed out');
      }
      throw error;
    }
  },

  async rateLimit(func, maxCalls = 20, timeWindowMs = 60000) {
    const timestamps = [];
    const now = new Date().getTime();
    
    // Clean up old timestamps
    while (timestamps.length && timestamps[0] < now - timeWindowMs) {
      timestamps.shift();
    }
    
    if (timestamps.length >= maxCalls) {
      const oldestCall = timestamps[0];
      const waitTime = timeWindowMs - (now - oldestCall);
      if (waitTime > 0) {
        this.delay(waitTime);
      }
    }
    
    timestamps.push(now);
    return await func();
  }
};