module.exports = {
  presets: [
    [
      '@babel/preset-env',
      {
        targets: '> 0.25%, not dead', // 可以根据需要调整目标浏览器
      },
    ],
  ],
};
