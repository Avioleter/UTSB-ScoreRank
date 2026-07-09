# 北科大教务系统 - 成绩排名增强 v2.4

油猴脚本，自动在成绩页面的「备注」列显示排名徽章，支持导出 Excel，计算加权平均分。

## 功能

- 🔵 **自动排名显示** — 每门课的「备注」列自动显示排名标签（`排名 3/31`），前三名金色高亮
- 📊 **加权平均分** — 右下角面板实时显示加权平均分
- 📥 **导出 Excel** — 一键导出成绩数据（含排名、学分、课程性质等 12 列）
- 🔄 **自动刷新** — 切换学期 / 翻页时自动重新注入排名

## 安装

### 方法一：一键安装（推荐）

1. 安装 [Tampermonkey](https://www.tampermonkey.net/) 浏览器扩展
2. 点击 [ustb-grades-enhancer.user.js](ustb-grades-enhancer.user.js) → 点击右上角 **Raw**（或直接打开 raw 链接），Tampermonkey 会自动弹出安装提示
3. 点击 **安装** 即可

### 方法二：手动复制

1. 安装 [Tampermonkey](https://www.tampermonkey.net/)
2. 新建脚本，复制 `ustb-grades-enhancer.user.js` 全部内容粘贴进去
3. `Ctrl+S` 保存

安装后打开 [成绩查询页面](https://byyt.ustb.edu.cn/cjgl/grcjcx/grcjcx)，等待约 1 秒即可看到排名。

## License

MIT

本项目仅用于个人学习交流，禁止大规模批量爬取、恶意访问教务系统，一切违规使用行为由使用者自行承担全部责任。
