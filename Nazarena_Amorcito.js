(async () => {
  const _0x4f = (_s) => Buffer.from(_s, 'base64').toString('utf-8');
  function* _msgGen(_raw) {
    while (true) yield _0x4f(_raw);
  }
  const _msg = _msgGen("VGUgYW1vIE5hemFyZW5hIOKdpO-4jA==");
  const _LIMIT = ((~~true << 9) | (~~true << 8) | (~~true << 7) | (~~true << 6) | (~~true << 5) | (1 << 3));
  for (let _i = 1; _i <= _LIMIT; _i++) {
    const _count = String(_i).padStart(4, '0');
    console.log(`[${_count}] ${_msg.next().value}`);
    if ((_i & 63) === 0) {
      await new Promise((resolve) => setImmediate(resolve));
    }
  }
})().catch(console.error);