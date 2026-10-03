/** 한 문서 열기가 시작되면 이전 비동기 열기의 결과 반영을 차단합니다. */
export function createDocumentOpenGate() {
  let current = null;
  return {
    begin() {
      const token = {};
      current = token;
      return {
        isCurrent: () => current === token,
        assertCurrent() {
          if (current === token) return;
          const error = new Error('다른 문서가 열려 이전 문서 열기를 취소했습니다.');
          error.code = 'DOCUMENT_OPEN_CANCELLED';
          throw error;
        },
      };
    },
    invalidate() { current = null; },
  };
}

export function isDocumentOpenCancelled(error) {
  return error?.code === 'DOCUMENT_OPEN_CANCELLED';
}
