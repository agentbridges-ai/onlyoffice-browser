export type OfficeLocale = 'zh-CN' | 'en-US';

export const OFFICE_LOCALE_STORAGE_KEY = 'onlyoffice-browser.locale';

export type OfficeCopy = {
  product: string;
  open: string;
  new: string;
  newWord: string;
  newSheet: string;
  newSlides: string;
  newDocumentName: string;
  noDocument: string;
  noDocumentHint: string;
  save: string;
  close: string;
  closeDialog: string;
  openDocuments: string;
  language: string;
  resources: string;
  resourcesReady: string;
  resourcesNeeded: string;
  resourcesUpdating: string;
  resourcesError: string;
  failedResources: string;
  repairFailedHint: string;
  resourceIntro: string;
  completePackage: string;
  completePackageHint: string;
  packageIncludes: string;
  downloadingAndVerifying: string;
  basicPreset: string;
  compatPreset: string;
  repair: string;
  allResources: string;
  advancedFonts: string;
  installed: string;
  download: string;
  remove: string;
  required: string;
  updateReady: string;
  updateNow: string;
  dirtyTitle: string;
  dirtyBody: string;
  discard: string;
  cancel: string;
  permission: string;
  authorize: string;
  error: string;
  version: string;
  word: string;
  cell: string;
  slide: string;
  core: string;
  fonts: string;
  officeDocument: string;
  officeDocuments: string;
  writePermissionError: string;
  planning: string;
  downloading: string;
  verifying: string;
  repairing: string;
  activating: string;
  readyToDownload: string;
  resourceStage: string;
  packageSegment: string;
  paused: string;
  pause: string;
  resume: string;
  retry: string;
  resourceProgress: string;
  resourceUpdateAvailable: string;
  resourceErrors: Record<
    'offline' | 'network' | 'timeout' | 'integrity' | 'quota' | 'manifest' | 'incompatible' | 'storage' | 'aborted',
    string
  >;
};

export const officeCopy: Record<OfficeLocale, OfficeCopy> = {
  'zh-CN': {
    product: 'OnlyOffice 浏览器版',
    open: '打开文件',
    new: '新建',
    newWord: 'Word 文档',
    newSheet: '电子表格',
    newSlides: '演示文稿',
    newDocumentName: '新建文档',
    noDocument: '打开或新建文档，开始在浏览器中编辑',
    noDocumentHint: '支持 DOCX、XLSX、PPTX、CSV 及常见旧版 Office 格式',
    save: '保存',
    close: '关闭',
    closeDialog: '关闭对话框',
    openDocuments: '已打开的文档',
    language: '语言',
    resources: 'Office 资源',
    resourcesReady: '资源已就绪',
    resourcesNeeded: '按需下载',
    resourcesUpdating: '正在更新资源',
    resourcesError: '资源需要修复',
    failedResources: '未能修复',
    repairFailedHint: '请重试；若仍失败，可能是在线资源暂时缺失。',
    resourceIntro: '组件、转换引擎和全部字体合并为一个完整 Office 资源包，一次安装即可使用。',
    completePackage: '完整 Office 资源包',
    completePackageHint: '一个版本、一个进度、一次完成',
    packageIncludes: '包含文字、表格、演示、转换组件和全部字体',
    downloadingAndVerifying: '正在下载并校验',
    basicPreset: '安装完整资源包',
    compatPreset: '安装 Office 兼容字体',
    repair: '检查并修复',
    allResources: '重新安装完整资源包',
    advancedFonts: '高级字体管理',
    installed: '已安装',
    download: '下载',
    remove: '移除',
    required: '基础字体',
    updateReady: '新版本已准备好。保存或关闭未保存的文档后将自动更新。',
    updateNow: '立即更新',
    dirtyTitle: '保存更改？',
    dirtyBody: '切换或关闭前，是否保存当前文档的更改？',
    discard: '不保存',
    cancel: '取消',
    permission: '需要重新授权此文件后才能恢复标签页。',
    authorize: '重新打开',
    error: '文档打开失败',
    version: '版本',
    word: '文字',
    cell: '表格',
    slide: '演示',
    core: '基础组件',
    fonts: '字体',
    officeDocument: 'Office 文档',
    officeDocuments: 'Office 文档',
    writePermissionError: '未获得文件写入权限。',
    planning: '正在准备下载',
    downloading: '正在下载',
    verifying: '正在校验',
    repairing: '正在检查并修复',
    activating: '正在完成更新',
    readyToDownload: '等待开始下载',
    resourceStage: '阶段',
    packageSegment: '资源分段',
    paused: '已暂停',
    pause: '暂停',
    resume: '继续',
    retry: '重试',
    resourceProgress: 'Office 资源进度',
    resourceUpdateAvailable: '有可用更新',
    resourceErrors: {
      offline: '网络已断开，联网后可继续',
      network: '资源下载失败',
      timeout: '资源下载超时',
      integrity: '资源完整性校验失败',
      quota: '浏览器存储空间不足',
      manifest: '资源清单无效',
      incompatible: '资源版本与当前应用不兼容',
      storage: '无法保存资源安装状态',
      aborted: '资源操作已取消',
    },
  },
  'en-US': {
    product: 'OnlyOffice Browser',
    open: 'Open files',
    new: 'New',
    newWord: 'Word document',
    newSheet: 'Spreadsheet',
    newSlides: 'Presentation',
    newDocumentName: 'New document',
    noDocument: 'Open or create a document to start editing in your browser',
    noDocumentHint: 'Supports DOCX, XLSX, PPTX, CSV, and common legacy Office formats',
    save: 'Save',
    close: 'Close',
    closeDialog: 'Close dialog',
    openDocuments: 'Open documents',
    language: 'Language',
    resources: 'Office resources',
    resourcesReady: 'Resources ready',
    resourcesNeeded: 'Downloads on demand',
    resourcesUpdating: 'Updating resources',
    resourcesError: 'Resources need repair',
    failedResources: 'Not repaired',
    repairFailedHint: 'Try again. If it still fails, an online resource may be temporarily unavailable.',
    resourceIntro: 'Components, conversion runtime, and every font are delivered as one complete Office package.',
    completePackage: 'Complete Office package',
    completePackageHint: 'One version, one progress bar, one installation',
    packageIncludes: 'Includes Word, Spreadsheet, Presentation, conversion components, and all fonts',
    downloadingAndVerifying: 'Downloading and verifying',
    basicPreset: 'Install complete package',
    compatPreset: 'Install Office-compatible fonts',
    repair: 'Check and repair',
    allResources: 'Reinstall complete package',
    advancedFonts: 'Advanced font management',
    installed: 'Installed',
    download: 'Download',
    remove: 'Remove',
    required: 'Essential',
    updateReady: 'A new version is ready. It will update automatically after unsaved documents are saved or closed.',
    updateNow: 'Update now',
    dirtyTitle: 'Save changes?',
    dirtyBody: 'Would you like to save changes before switching or closing this document?',
    discard: 'Discard',
    cancel: 'Cancel',
    permission: 'Authorize this file again to restore the tab.',
    authorize: 'Open again',
    error: 'Unable to open document',
    version: 'Version',
    word: 'Word',
    cell: 'Spreadsheet',
    slide: 'Presentation',
    core: 'Essentials',
    fonts: 'Fonts',
    officeDocument: 'Office document',
    officeDocuments: 'Office documents',
    writePermissionError: 'Write permission was not granted.',
    planning: 'Preparing download',
    downloading: 'Downloading',
    verifying: 'Verifying',
    repairing: 'Checking and repairing',
    activating: 'Activating resources',
    readyToDownload: 'Ready to download',
    resourceStage: 'Stage',
    packageSegment: 'Package segment',
    paused: 'Paused',
    pause: 'Pause',
    resume: 'Resume',
    retry: 'Retry',
    resourceProgress: 'Office resource progress',
    resourceUpdateAvailable: 'Update available',
    resourceErrors: {
      offline: 'Offline. The download can resume when the connection returns.',
      network: 'A resource download failed.',
      timeout: 'A resource download timed out.',
      integrity: 'A resource failed integrity verification.',
      quota: 'There is not enough browser storage.',
      manifest: 'The resource manifest is invalid.',
      incompatible: 'The resource release is incompatible with this app.',
      storage: 'The installation state could not be saved.',
      aborted: 'The resource operation was cancelled.',
    },
  },
};

export function normalizeOfficeLocale(value: string | null | undefined): OfficeLocale | null {
  if (!value) return null;
  const normalized = value.toLowerCase();
  if (normalized.startsWith('zh')) return 'zh-CN';
  if (normalized.startsWith('en')) return 'en-US';
  return null;
}

export function resolveOfficeLocale(
  storedLocale: string | null | undefined,
  browserLanguages: readonly string[],
): OfficeLocale {
  const stored = normalizeOfficeLocale(storedLocale);
  if (stored) return stored;
  for (const language of browserLanguages) {
    const locale = normalizeOfficeLocale(language);
    if (locale) return locale;
  }
  return 'en-US';
}
