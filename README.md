# mac-slideSCI

PowerPoint for Mac 的 SCI 演示辅助插件：图片排版、对齐均分、Panel label/caption、字体统一、可编辑组合库。

## 在线托管（GitHub Pages）

- 任务窗格：https://cherryamme.github.io/mac-slideSCI/src/taskpane/index.html
- 命令文件：https://cherryamme.github.io/mac-slideSCI/src/commands/commands.html
- 生产 manifest：https://cherryamme.github.io/mac-slideSCI/manifest.xml

每次推送到 `main` 后，GitHub Actions 会自动构建并发布到 GitHub Pages。

## 用户安装（Mac PowerPoint）

1. 下载 manifest：

   ```bash
   curl -L https://cherryamme.github.io/mac-slideSCI/manifest.xml \
     -o "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/mac-slideSCI.xml"
   ```

   也可以手动下载后放到该目录（如目录不存在，请先 `mkdir -p`）。

2. 重启 PowerPoint for Mac。

3. 顶部 ribbon 会出现 `mac-slideSCI` 选项卡，包含：排版、对齐、均分、图注、布局设置、位置尺寸、标注、字体、组合、设置等分组。

## 开发

```bash
npm install
npm run dev          # 启动本地 HTTPS dev server (https://localhost:3000)
npm run sideload:mac # 把本地 manifest.xml 复制到 PowerPoint 侧载目录
```

其他脚本：

- `npm test` 运行测试
- `npm run typecheck` 类型检查
- `npm run validate:manifest` 校验 manifest
- `npm run build` 生产构建（本地用）
- `PUBLIC_BASE_URL=https://cherryamme.github.io/mac-slideSCI PUBLIC_BASE_PATH=/mac-slideSCI/ npm run build:pages` 构建发布用产物到 `dist/`

## GitHub Pages 一次性设置

1. 仓库 **Settings → Pages → Source** 选 `GitHub Actions`。
2. 推送到 `main`，Actions 会自动跑 `Deploy to GitHub Pages` 工作流。
3. 第一次部署完成后即可访问上面的 URL。
