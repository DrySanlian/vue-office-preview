function createAbortError(){
    const error = new Error('Excel解析已取消');
    error.name = 'AbortError';
    return error;
}

export function getWorkerOptions(options = {}){
    return {
        xls: !!options.xls,
        minColLength: options.minColLength,
        minRowLength: options.minRowLength,
        widthOffset: options.widthOffset,
        heightOffset: options.heightOffset
    };
}

export class ExcelWorkerClient {
    constructor(createWorker){
        this.createWorker = createWorker;
        this.worker = null;
        this.pending = null;
    }

    parse(buffer, options){
        this.cancel();
        const worker = this.createWorker();
        this.worker = worker;
        return new Promise((resolve, reject) => {
            let settled = false;
            const finish = (callback, value) => {
                if (settled) {
                    return;
                }
                settled = true;
                worker.terminate();
                if (this.worker === worker) {
                    this.worker = null;
                    this.pending = null;
                }
                callback(value);
            };
            this.pending = {
                reject: error => finish(reject, error)
            };
            worker.onmessage = event => {
                const data = event.data || {};
                if (data.type === 'success') {
                    finish(resolve, data);
                } else {
                    const error = new Error(data.error?.message || 'Excel解析失败');
                    error.name = data.error?.name || 'Error';
                    error.stack = data.error?.stack || error.stack;
                    finish(reject, error);
                }
            };
            worker.onerror = event => {
                finish(reject, event.error || new Error(event.message || 'Excel Worker加载失败'));
            };
            try {
                worker.postMessage({type: 'parse', buffer, options}, [buffer]);
            } catch (error) {
                finish(reject, error);
            }
        });
    }

    cancel(){
        const pending = this.pending;
        const worker = this.worker;
        this.pending = null;
        this.worker = null;
        if (worker) {
            worker.terminate();
        }
        if (pending) {
            pending.reject(createAbortError());
        }
    }

    destroy(){
        this.cancel();
        this.createWorker = null;
    }
}
