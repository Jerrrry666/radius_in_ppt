/* One host operation at a time, including its final implicit sync. */
(function (root) {
  function createHostQueue(runContext) {
    let tail = Promise.resolve();
    let pending = 0;
    return {
      get busy() { return pending > 0; },
      run(callback, isCurrent) {
        pending++;
        const work = tail.then(() => {
          if (isCurrent && !isCurrent()) return undefined;
          return runContext(callback);
        });
        const result = work.finally(() => { pending--; });
        tail = result.catch(() => {});
        return result;
      },
    };
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = { createHostQueue };
  root.HostQueue = { createHostQueue };
})(typeof window !== 'undefined' ? window : globalThis);
