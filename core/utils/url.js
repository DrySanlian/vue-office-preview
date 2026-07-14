const noop = () => {};

export function getUrlInfo(src, options){
    if(typeof src === 'string'){
        return {url: src, revoke: noop};
    }else if(src instanceof Blob){
        const url = URL.createObjectURL(src);
        return {url, revoke: () => URL.revokeObjectURL(url)};
    }else if(src instanceof ArrayBuffer){
        const url = URL.createObjectURL(new Blob([src], options));
        return {url, revoke: () => URL.revokeObjectURL(url)};
    }else{
        return {url: src, revoke: noop};
    }
}

export function getUrl(src, options){
    return getUrlInfo(src, options).url;
}

export function loadScript(src){
    return new Promise(((resolve, reject) => {
        let script = document.createElement('script');
        script.src = src;
        script.onload = function (){
            resolve();
        };
        script.onerror = function (){
            reject();
        };
        document.body.append(script);
    }));
}

export async function download(filename, data){
    if(!data){
        return; 
    }
   if (data instanceof ArrayBuffer) {
       data = new Blob([data]);
    }
    const url = URL.createObjectURL(data);
    downloadFile(filename, url);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function downloadFile(filename, href){
    let eleLink = document.createElement('a');
    eleLink.download = filename;
    eleLink.style.display = 'none';
    eleLink.href = href;
    document.body.appendChild(eleLink);
    eleLink.click();
    document.body.removeChild(eleLink);
}
