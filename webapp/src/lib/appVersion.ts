import versionInfo from '../../version.json'

/** 화면 푸터에 표시하는 버전 문구. 기준은 webapp/version.json 하나이며 exe 실행 창도 같은 값을 쓴다. */
export const APP_VERSION_LABEL = `v${versionInfo.version} (${versionInfo.date})`
