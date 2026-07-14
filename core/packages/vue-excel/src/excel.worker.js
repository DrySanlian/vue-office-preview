import {readExcelData, transferExcelToSpreadSheet} from './excel';

function toTransferableBuffer(buffer){
    if (buffer instanceof ArrayBuffer) {
        return buffer;
    }
    if (ArrayBuffer.isView(buffer)) {
        return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    }
    return null;
}

self.onmessage = async event => {
    const {buffer, options = {}} = event.data || {};
    try {
        const workbook = await readExcelData(buffer, options.xls);
        if (!workbook._worksheets || workbook._worksheets.length === 0) {
            throw new Error('未获取到数据，可能文件格式不正确或文件已损坏');
        }
        const result = transferExcelToSpreadSheet(workbook, options);
        const transferables = [];
        const sourceBuffer = toTransferableBuffer(buffer);
        if (sourceBuffer) {
            transferables.push(sourceBuffer);
            result.buffer = sourceBuffer;
        }
        result.medias = result.medias.map(media => {
            const mediaBuffer = toTransferableBuffer(media.buffer);
            if (mediaBuffer) {
                transferables.push(mediaBuffer);
                return {...media, buffer: mediaBuffer};
            }
            return media;
        });
        self.postMessage({type: 'success', ...result}, transferables);
    } catch (error) {
        const data = {
            type: 'error',
            error: {
                name: error.name,
                message: error.message || String(error),
                stack: error.stack
            }
        };
        const sourceBuffer = toTransferableBuffer(buffer);
        if (sourceBuffer) {
            data.buffer = sourceBuffer;
            self.postMessage(data, [sourceBuffer]);
        } else {
            self.postMessage(data);
        }
    }
};
