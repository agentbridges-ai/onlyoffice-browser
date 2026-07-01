(() => {
  'use strict';

  class AsyncMessageHub {
    handlers = new Map();
    counter = 0;
    queue = [];
    debug;

    constructor(debug) {
      this.debug = debug;
    }

    addHandler(handler) {
      const wrapped = this.debug
        ? (message) => {
            console.log(this.debug, message);
            handler(message);
          }
        : handler;
      const id = this.counter;
      this.counter += 1;
      this.handlers.set(id, wrapped);
      if (this.queue.length > 0) {
        for (const message of this.queue) this.callHandlers(message);
        this.queue = [];
      }
      return new HandlerRegistration(this.handlers, id);
    }

    fire(message) {
      if (this.handlers.size > 0) {
        this.callHandlers(message);
      } else {
        this.queue.push(structuredClone(message));
      }
    }

    callHandlers(message) {
      for (const handler of this.handlers.values()) {
        const cloned = structuredClone(message);
        setTimeout(() => handler(cloned));
      }
    }
  }

  class HandlerRegistration {
    handlers;
    id;

    constructor(handlers, id) {
      this.handlers = handlers;
      this.id = id;
    }

    remove() {
      this.handlers.delete(this.id);
    }
  }

  function mergeConfig(base, overrides) {
    if (typeof base !== 'object') return overrides;
    const result = Object.assign({}, base);
    for (const key of Object.keys(overrides)) {
      result[key] = mergeConfig(base[key], overrides[key]);
    }
    return result;
  }

  function noop() {}

  function createDeferredEvent(fireOnce = false) {
    const handlers = [];
    let fired = false;
    return {
      reg(handler) {
        if (fireOnce && fired) {
          setTimeout(handler);
        } else {
          handlers.push(handler);
        }
      },
      unreg(handler) {
        const index = handlers.indexOf(handler);
        if (index !== -1) {
          handlers.splice(index, 1);
        } else {
          console.log('event handler was already unregistered');
        }
      },
      fire(...args) {
        if (fireOnce && fired) return;
        fired = true;
        handlers.forEach((handler) => {
          handler(...args);
        });
      },
    };
  }

  let OriginalDocEditor;

  class BrowserDocEditor {
    waitForAppReady;
    origEditor;
    fromOOHandler = new AsyncMessageHub();
    toOOHandler = new AsyncMessageHub();
    placeholderId;
    server;
    fromOOHandle;
    corruptionWarningHandler = new AsyncMessageHub();

    constructor(placeholderId, config) {
      this.placeholderId = placeholderId;
      this.init(config).catch((error) => {
        console.error(error);
      });
    }

    async init(config) {
      let resolveAppReady;
      await loadOriginalApi;
      this.waitForAppReady = new Promise((resolve) => {
        resolveAppReady = resolve;
      });
      this.waitForAppReady.then(config?.events?.onAppReady ?? noop).catch(noop);

      const merged = mergeConfig(config, {
        events: {
          onAppReady: resolveAppReady,
          cryptPadSendMessageFromOO: (event) => {
            this.fromOOHandler.fire(event.data.msg);
          },
          cryptPadCorruptionWarningHandler: (event) => {
            this.corruptionWarningHandler.fire(event.data.duplicateId);
          },
        },
      });

      this.origEditor = new OriginalDocEditor(this.placeholderId, merged);
      this.toOOHandler.addHandler((message) => this.origEditor.cryptPadMessageToOO(message));
    }

    installLegacyChannel() {
      const editorMessages = createDeferredEvent();
      const frameWindow = this.getIframe().contentWindow;
      window.addEventListener('message', (event) => {
        if (event.source === frameWindow) editorMessages.fire(event);
      });

      installRpcBridge(
        editorMessages,
        (message) => {
          frameWindow.postMessage(message);
        },
        (rpc) => {
          this.toOOHandler.addHandler((message) => {
            rpc.event('CMD', message);
          });
          rpc.on('CMD', (message) => {
            this.fromOOHandler.fire(message);
          });
        },
      );
    }

    destroyEditor() {
      this.fromOOHandle?.remove();
      this.origEditor.destroyEditor();
    }

    getIframe() {
      return document.querySelector('iframe[name="frameEditor"]');
    }

    injectCSS(css) {
      const head = this.getIframe().contentDocument.querySelector('head');
      const style = document.createElement('style');
      style.innerText = css;
      head.appendChild(style);
    }

    sendMessageToOO(message) {
      this.toOOHandler.fire(message);
    }

    connectMockServer(server) {
      this.server = server;
      window.APP.getImageURL = server.getImageURL
        ? (name, callback) => {
            server.getImageURL(name).then(callback).catch((error) => console.error(error));
          }
        : (_name, callback) => callback('');

      this.fromOOHandle = this.fromOOHandler.addHandler((message) => {
        if (message?.type === 'auth') {
          this.handleAuth(message);
        }
        server.handleMessage?.(message, (response) => this.sendMessageToOO(response));
        server.onMessage(message);
      });
      if (server.onCorruptionWarning) {
        this.corruptionWarningHandler.addHandler(server.onCorruptionWarning);
      }
    }

    handleAuth(message) {
      const changes = this.server.getInitialChanges ? this.server.getInitialChanges() : [];
      const participants = this.server.getParticipants();
      if (changes.length > 0) {
        this.sendMessageToOO({ type: 'authChanges', changes });
      }
      this.sendMessageToOO({
        type: 'auth',
        result: 1,
        sessionId: 'session-id',
        participants: participants.list,
        locks: [],
        changes,
        changesIndex: 0,
        indexUser: participants.index,
        buildVersion: this.server.buildVersion || '9.3.0',
        buildNumber: this.server.buildNumber || 140,
        licenseType: 3,
      });
      this.sendMessageToOO({
        type: 'documentOpen',
        data: {
          type: 'open',
          status: 'ok',
          data: { 'Editor.bin': message.openCmd.url },
        },
      });
      this.server.onAuth?.();
    }

    serviceCommand(...args) {
      return this.origEditor.serviceCommand(...args);
    }
    showMessage(...args) {
      return this.origEditor.showMessage(...args);
    }
    processSaveResult(...args) {
      return this.origEditor.processSaveResult(...args);
    }
    processRightsChange(...args) {
      return this.origEditor.processRightsChange(...args);
    }
    denyEditingRights(...args) {
      return this.origEditor.denyEditingRights(...args);
    }
    refreshHistory(...args) {
      return this.origEditor.refreshHistory(...args);
    }
    setHistoryData(...args) {
      return this.origEditor.setHistoryData(...args);
    }
    setEmailAddresses(...args) {
      return this.origEditor.setEmailAddresses(...args);
    }
    setActionLink(...args) {
      return this.origEditor.setActionLink(...args);
    }
    processMailMerge(...args) {
      return this.origEditor.processMailMerge(...args);
    }
    downloadAs(...args) {
      return this.origEditor.downloadAs(...args);
    }
    getEditorApi() {
      const app = this.getIframe()?.contentWindow?.__onlyOfficeBrowserApplication;
      return app?.getController?.('Main')?.api;
    }
    getEditorWindow() {
      return this.getIframe()?.contentWindow || null;
    }
    asc_nativeGetFile3(...args) {
      const api = this.getEditorApi();
      if (!api || typeof api.asc_nativeGetFile3 !== 'function') {
        throw new Error('OnlyOffice native file API is not ready');
      }
      return api.asc_nativeGetFile3(...args);
    }
    asc_nativeGetPDF(...args) {
      const api = this.getEditorApi();
      if (!api || typeof api.asc_nativeGetPDF !== 'function') {
        throw new Error('OnlyOffice native PDF API is not ready');
      }
      return api.asc_nativeGetPDF(...args);
    }
    attachMouseEvents(...args) {
      return this.origEditor.attachMouseEvents(...args);
    }
    detachMouseEvents(...args) {
      return this.origEditor.detachMouseEvents(...args);
    }
    setUsers(...args) {
      return this.origEditor.setUsers(...args);
    }
    showSharingSettings(...args) {
      return this.origEditor.showSharingSettings(...args);
    }
    setSharingSettings(...args) {
      return this.origEditor.setSharingSettings(...args);
    }
    insertImage(...args) {
      return this.origEditor.insertImage(...args);
    }
    setMailMergeRecipients(...args) {
      return this.origEditor.setMailMergeRecipients(...args);
    }
    setRevisedFile(...args) {
      return this.origEditor.setRevisedFile(...args);
    }
    setFavorite(...args) {
      return this.origEditor.setFavorite(...args);
    }
    requestClose(...args) {
      return this.origEditor.requestClose(...args);
    }
    grabFocus(...args) {
      return this.origEditor.grabFocus(...args);
    }
    blurFocus(...args) {
      return this.origEditor.blurFocus(...args);
    }
    setReferenceData(...args) {
      return this.origEditor.setReferenceData(...args);
    }
  }

  function installRpcBridge(editorMessages, postToFrame, setup) {
    let ready = false;
    const backlog = [];
    const readyEvent = createDeferredEvent(true);

    editorMessages.reg((event) => {
      if (ready) return;
      const data = event.data;
      if (data === '_READY') {
        postToFrame('_READY');
        ready = true;
        readyEvent.fire();
        backlog.forEach((backlogged) => {
          editorMessages.fire(backlogged);
        });
        return;
      }
      backlog.push(data);
    });

    const handlers = {};
    const pendingResponses = {};
    const pendingAcks = {};
    const registeredEvents = [];
    const whenRegistered = {};

    const rpc = {
      query(name, content, callback, options) {
        const txid = Math.random().toString(16).replace('0.', '') + Math.random().toString(16).replace('0.', '');
        const actualOptions = options || {};
        const timeout = actualOptions.timeout || 30000;
        let timer;
        if (timeout > 0) {
          timer = setTimeout(() => {
            delete pendingResponses[txid];
            callback('TIMEOUT');
          }, timeout);
        }
        pendingAcks[txid] = (unhandled) => {
          clearTimeout(timer);
          delete pendingAcks[txid];
          if (unhandled) {
            delete pendingResponses[txid];
            callback('UNHANDLED');
          }
        };
        pendingResponses[txid] = (message, event) => {
          delete pendingResponses[txid];
          callback(undefined, message.content, event);
        };
        readyEvent.reg(() => {
          const message = { txid, content, q: name, raw: actualOptions.raw };
          postToFrame(actualOptions.raw ? message : JSON.stringify(message));
        });
      },
    };

    const event = (rpc.event = (name, content, options) => {
      const actualOptions = options || {};
      readyEvent.reg(() => {
        const message = { content, q: name, raw: actualOptions.raw };
        postToFrame(actualOptions.raw ? message : JSON.stringify(message));
      });
    });

    rpc.on = (name, handler, raw) => {
      const wrapped = (message, callback, isRaw) => {
        handler(
          message.content,
          (response) => {
            const reply = { txid: message.txid, content: response };
            postToFrame(isRaw ? reply : JSON.stringify(reply));
          },
          raw,
        );
      };
      (handlers[name] = handlers[name] || []).push(wrapped);
      if (!raw) event('EV_REGISTER_HANDLER', name);
      return {
        stop() {
          const index = handlers[name].indexOf(wrapped);
          if (index !== -1) handlers[name].splice(index, 1);
        },
      };
    };

    rpc.whenReg = (name, callback, retain) => {
      let shouldRetain = retain;
      if (registeredEvents.indexOf(name) > -1) {
        callback();
      } else {
        shouldRetain = true;
      }
      if (shouldRetain) {
        (whenRegistered[name] = whenRegistered[name] || []).push(callback);
      }
    };

    rpc.onReg = (name, callback) => {
      rpc.whenReg(name, callback, true);
    };

    rpc.on('EV_REGISTER_HANDLER', (name) => {
      if (whenRegistered[name]) {
        whenRegistered[name].forEach((callback) => {
          callback();
        });
        delete whenRegistered[name];
      }
      registeredEvents.push(name);
    });

    let rpcReady = false;
    rpc.onReady = (callback) => {
      if (rpcReady) {
        callback();
      } else if (typeof callback === 'function') {
        rpc.on('EV_RPC_READY', () => {
          rpcReady = true;
          callback();
        });
      }
    };

    rpc.ready = () => {
      rpc.whenReg('EV_RPC_READY', () => {
        rpc.event('EV_RPC_READY');
      });
    };

    editorMessages.reg((eventMessage) => {
      if (!ready) return;
      if (!eventMessage.data || eventMessage.data === '_READY') return;
      let parsed;
      try {
        parsed = typeof eventMessage.data === 'object' ? eventMessage.data : JSON.parse(eventMessage.data);
      } catch (error) {
        console.warn(error);
        return;
      }

      if (parsed.ack !== undefined) {
        pendingAcks[parsed.txid]?.(!parsed.ack);
      } else if (typeof parsed.q === 'string') {
        if (handlers[parsed.q]) {
          if (parsed.txid) postToFrame(JSON.stringify({ txid: parsed.txid, ack: true }));
          handlers[parsed.q].forEach((handler) => {
            handler(parsed || JSON.parse(eventMessage.data), eventMessage, parsed?.raw);
            parsed = undefined;
          });
        } else if (parsed.txid) {
          postToFrame(JSON.stringify({ txid: parsed.txid, ack: false }));
        }
      } else if (parsed.q === undefined && pendingResponses[parsed.txid]) {
        pendingResponses[parsed.txid](parsed, eventMessage);
      }
    });

    postToFrame('_READY');
    setup(rpc);
  }

  const loadOriginalApi = (async function () {
    let apiUrl;
    let apiScript;
    for (const script of document.getElementsByTagName('script')) {
      try {
        if (new URL(script.src).pathname.endsWith('web-apps/apps/api/documents/api.js')) {
          apiUrl = script.src;
          apiScript = script;
          break;
        }
      } catch (error) {
        if (!(error instanceof TypeError)) throw error;
      }
    }

    const script = document.createElement('script');
    script.setAttribute('type', 'text/javascript');
    const originalUrl = new URL('api-orig.js', apiUrl);
    originalUrl.search = new URL(apiUrl).search;
    script.setAttribute('src', originalUrl.href);
    const loaded = new Promise((resolve) => {
      script.addEventListener('load', () => resolve(), { once: true });
    });

    apiScript.after(script);
    const runtimeWindow = window;
    runtimeWindow.DocsAPI = runtimeWindow.DocsAPI ?? {};
    runtimeWindow.DocsAPI.DocEditor = BrowserDocEditor;
    await loaded;
    OriginalDocEditor = runtimeWindow.DocsAPI.DocEditor;
    runtimeWindow.DocsAPI.DocEditor = BrowserDocEditor;
  })();
})();
