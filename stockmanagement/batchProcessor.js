// batchProcessor.js
// Version 1.0
// 分割処理による確実なデータ更新システム

var BatchProcessor = {
  
  // 設定
  CONFIG: {
    MAX_BATCH_SIZE: 5,           // 1バッチあたりの最大処理件数
    SAFE_TIME_MARGIN: 4.5 * 60 * 1000,  // 安全マージン（4.5分）
    BATCH_INTERVAL: 2000,        // バッチ間の待機時間（ミリ秒）
    ITEM_INTERVAL: 500,          // アイテム間の待機時間（ミリ秒）
    MAX_RETRIES: 3,              // 最大リトライ回数
    PROGRESS_SAVE_INTERVAL: 10   // 進捗保存間隔
  },

  /**
   * 分割処理による確実なデータ更新
   * @param {Array} items 処理対象アイテム
   * @param {Function} processor 処理関数
   * @param {Object} options オプション
   * @returns {Object} 処理結果
   */
  processInBatches: function(items, processor, options = {}) {
    const config = { ...this.CONFIG, ...options };
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // 処理状態の初期化
    const processId = options.processId || this.generateProcessId();
    const startTime = new Date();
    
    // 前回の処理状態を復元
    const savedState = this.loadProcessState(processId);
    const startIndex = savedState ? savedState.lastProcessedIndex + 1 : 0;
    
    LoggerManager.info(`分割処理を開始します: ${items.length}件中${startIndex}番目から`, {
      processId: processId,
      totalItems: items.length,
      startIndex: startIndex,
      batchSize: config.MAX_BATCH_SIZE
    });

    // タイムアウト管理開始
    TimeoutManager.startExecutionTimer();
    
    let processedCount = startIndex;
    let successCount = 0;
    let errorCount = 0;
    let currentBatch = Math.floor(startIndex / config.MAX_BATCH_SIZE) + 1;
    const totalBatches = Math.ceil(items.length / config.MAX_BATCH_SIZE);
    
    try {
      // バッチ処理ループ
      for (let i = startIndex; i < items.length; i += config.MAX_BATCH_SIZE) {
        const batchStart = i;
        const batchEnd = Math.min(i + config.MAX_BATCH_SIZE, items.length);
        const batchItems = items.slice(batchStart, batchEnd);
        
        // キャンセルチェック
        if (scriptProperties.getProperty('cancelFlag') === 'true') {
          LoggerManager.info('処理がキャンセルされました');
          this.saveProcessState(processId, {
            lastProcessedIndex: i - 1,
            processedCount: processedCount,
            successCount: successCount,
            errorCount: errorCount,
            status: 'cancelled'
          });
          return { status: 'cancelled', processedCount, successCount, errorCount };
        }
        
        // タイムアウトチェック
        if (TimeoutManager.isExecutionTimedOut()) {
          LoggerManager.warn(`タイムアウトのため処理を中断: ${processedCount}/${items.length}件完了`);
          this.saveProcessState(processId, {
            lastProcessedIndex: i - 1,
            processedCount: processedCount,
            successCount: successCount,
            errorCount: errorCount,
            status: 'timeout'
          });
          return { status: 'timeout', processedCount, successCount, errorCount };
        }
        
        // 安全マージンチェック
        const elapsedTime = new Date() - startTime;
        if (elapsedTime > config.SAFE_TIME_MARGIN) {
          LoggerManager.warn(`安全マージンのため処理を中断: ${processedCount}/${items.length}件完了`);
          this.saveProcessState(processId, {
            lastProcessedIndex: i - 1,
            processedCount: processedCount,
            successCount: successCount,
            errorCount: errorCount,
            status: 'safe_timeout'
          });
          return { status: 'safe_timeout', processedCount, successCount, errorCount };
        }
        
        // バッチ処理開始
        LoggerManager.info(`バッチ ${currentBatch}/${totalBatches} を処理中: ${batchItems.length}件`, {
          batchStart: batchStart,
          batchEnd: batchEnd,
          processedCount: processedCount,
          totalItems: items.length
        });
        
        // バッチ内の各アイテムを処理
        for (let j = 0; j < batchItems.length; j++) {
          const item = batchItems[j];
          const itemIndex = batchStart + j;
          
          try {
            // アイテム処理
            const result = processor(item, itemIndex, items);
            if (result && result.success !== false) {
              successCount++;
            } else {
              errorCount++;
              LoggerManager.warn(`アイテム処理失敗: ${itemIndex}`, result);
            }
          } catch (error) {
            errorCount++;
            LoggerManager.error(`アイテム処理エラー: ${itemIndex}`, error);
          }
          
          processedCount++;
          
          // 進捗更新
          if (processedCount % config.PROGRESS_SAVE_INTERVAL === 0) {
            this.updateProgress(processedCount, items.length, currentBatch, totalBatches);
            this.saveProcessState(processId, {
              lastProcessedIndex: itemIndex,
              processedCount: processedCount,
              successCount: successCount,
              errorCount: errorCount,
              status: 'processing'
            });
          }
          
          // アイテム間待機
          if (j < batchItems.length - 1) {
            Utilities.sleep(config.ITEM_INTERVAL);
          }
        }
        
        // バッチ完了
        LoggerManager.info(`バッチ ${currentBatch}/${totalBatches} 完了: ${batchItems.length}件処理`, {
          successCount: successCount,
          errorCount: errorCount,
          processedCount: processedCount
        });
        
        // 進捗保存
        this.saveProcessState(processId, {
          lastProcessedIndex: batchEnd - 1,
          processedCount: processedCount,
          successCount: successCount,
          errorCount: errorCount,
          status: 'processing'
        });
        
        currentBatch++;
        
        // バッチ間待機（最後のバッチ以外）
        if (batchEnd < items.length) {
          Utilities.sleep(config.BATCH_INTERVAL);
        }
      }
      
      // 処理完了
      LoggerManager.info(`分割処理完了: ${processedCount}/${items.length}件処理`, {
        successCount: successCount,
        errorCount: errorCount,
        totalTime: new Date() - startTime
      });
      
      // 完了状態を保存
      this.saveProcessState(processId, {
        lastProcessedIndex: items.length - 1,
        processedCount: processedCount,
        successCount: successCount,
        errorCount: errorCount,
        status: 'completed'
      });
      
      return { 
        status: 'completed', 
        processedCount, 
        successCount, 
        errorCount,
        totalTime: new Date() - startTime
      };
      
    } catch (error) {
      LoggerManager.error('分割処理中にエラーが発生しました', error);
      this.saveProcessState(processId, {
        lastProcessedIndex: processedCount - 1,
        processedCount: processedCount,
        successCount: successCount,
        errorCount: errorCount,
        status: 'error',
        error: error.message
      });
      throw error;
    } finally {
      TimeoutManager.cleanup();
    }
  },

  /**
   * 処理状態を保存
   */
  saveProcessState: function(processId, state) {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty(`batch_state_${processId}`, JSON.stringify({
        ...state,
        timestamp: new Date().toISOString()
      }));
    } catch (error) {
      LoggerManager.error('処理状態の保存に失敗しました', error);
    }
  },

  /**
   * 処理状態を読み込み
   */
  loadProcessState: function(processId) {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const stateJson = scriptProperties.getProperty(`batch_state_${processId}`);
      return stateJson ? JSON.parse(stateJson) : null;
    } catch (error) {
      LoggerManager.error('処理状態の読み込みに失敗しました', error);
      return null;
    }
  },

  /**
   * 処理状態をクリア
   */
  clearProcessState: function(processId) {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.deleteProperty(`batch_state_${processId}`);
      LoggerManager.debug(`処理状態をクリアしました: ${processId}`);
    } catch (error) {
      LoggerManager.error('処理状態のクリアに失敗しました', error);
    }
  },

  /**
   * プロセスIDを生成
   */
  generateProcessId: function() {
    return `batch_${new Date().getTime()}_${Math.random().toString(36).substr(2, 9)}`;
  },

  /**
   * 固定のプロセスIDを生成（継続処理用）
   */
  generateFixedProcessId: function(processType) {
    return `batch_${processType}_fixed`;
  },

  /**
   * 進捗更新
   */
  updateProgress: function(processed, total, currentBatch, totalBatches) {
    const percent = Math.round((processed / total) * 100);
    UIManager.updateProgress(processed, total, 
      `バッチ ${currentBatch}/${totalBatches} 処理中: ${processed}/${total} (${percent}%)`, {
        phase: 'バッチ処理',
        currentBatch: currentBatch,
        totalBatches: totalBatches,
        percentComplete: percent
      }
    );
  },

  /**
   * 分割処理の継続実行
   */
  continueProcessing: function(processId, items, processor, options = {}) {
    const savedState = this.loadProcessState(processId);
    if (!savedState) {
      throw new Error('継続処理の状態が見つかりません');
    }
    
    LoggerManager.info(`分割処理を継続します: ${processId}`, savedState);
    
    // 残りのアイテムを処理
    const remainingItems = items.slice(savedState.lastProcessedIndex + 1);
    return this.processInBatches(remainingItems, processor, options);
  },

  /**
   * 最適化されたバッチサイズを計算
   */
  calculateOptimalBatchSize: function(totalItems, complexity = 5) {
    // 複雑さに基づいて基本バッチサイズを調整
    const baseSize = Math.max(1, Math.floor(10 / complexity));
    
    // データ量に基づいて調整
    if (totalItems > 100) {
      return Math.min(5, baseSize);
    } else if (totalItems > 50) {
      return Math.min(8, baseSize * 1.5);
    } else {
      return Math.min(totalItems, Math.max(3, baseSize));
    }
  }
}; 