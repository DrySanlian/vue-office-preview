import hack from './hack.js';
import docx from '../../vue-docx/src/docx';
import {download as downloadFile} from '../../../utils/url.js';
class JsDocxPreview {
    container = null;
    wrapper = null;
    wrapperMain = null;
    options = {};
    requestOptions = {};
    fileData = null;
    renderVersion = 0;
    destroyed = false;
    requestController = null;
    
    constructor(container, options={}, requestOptions={}) {
        this.container = container;
        this.options = options;
        this.requestOptions = requestOptions;
        this.createWrapper();
    }
    createWrapper(){
        this.wrapper = document.createElement('div');
        this.wrapper.className = 'vue-office-docx';
        this.wrapperMain = document.createElement('div');
        this.wrapperMain.className = 'vue-office-docx-main';
        this.wrapper.appendChild(this.wrapperMain);
        this.container.appendChild(this.wrapper);
    }

    setOptions(options) {
        this.options = options;
    }
    setRequestOptions(requestOptions) {
        this.requestOptions = requestOptions;
    }
    preview(src){
        const version = ++this.renderVersion;
        this.cancelRequest();
        const controller = typeof AbortController === 'function' ? new AbortController() : null;
        this.requestController = controller;
        const requestOptions = controller
            ? {...this.requestOptions, signal: controller.signal}
            : this.requestOptions;
        return new Promise((resolve, reject) => {
            docx.getData(src, requestOptions).then(async res =>{
                if(this.requestController === controller){
                    this.requestController = null;
                }
                if(this.destroyed || version !== this.renderVersion){
                    resolve();
                    return;
                }
                this.fileData = await docx.getBlob(res);
                if(this.destroyed || version !== this.renderVersion){
                    resolve();
                    return;
                }
                docx.render(this.fileData, this.wrapperMain, this.options).then(() => {
                    if(this.destroyed || version !== this.renderVersion){
                        resolve();
                        return;
                    }
                    resolve();
                }).catch(e => {
                    if(this.destroyed || version !== this.renderVersion){
                        resolve();
                        return;
                    }
                    this.wrapperMain.innerHTML = '';
                    reject(e);
                });
            }).catch(err=>{
                if(this.requestController === controller){
                    this.requestController = null;
                }
                if(this.destroyed || version !== this.renderVersion){
                    resolve();
                    return;
                }
                this.wrapperMain.innerHTML = '';
                reject(err);
            });
        });
    }
    save(fileName){
        downloadFile(fileName || `js-preview-docx-${new Date().getTime()}.docx`,this.fileData);
    }
    destroy(){
        if(this.destroyed){
            return;
        }
        this.destroyed = true;
        this.renderVersion += 1;
        this.cancelRequest();
        this.container.removeChild(this.wrapper);
        this.container = null;
        this.wrapper = null;
        this.wrapperMain = null;
        this.options = null;
        this.requestOptions = null;
    }
    cancelRequest(){
        if(this.requestController){
            this.requestController.abort();
            this.requestController = null;
        }
    }
}
export function init(container, options, requestOptions){
    return new JsDocxPreview(container, options, requestOptions);
}
