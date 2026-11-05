# 眼镜零售门店促销 Memo 自动生成器 (OO / SGH / RB / LC)

## 1. 包含文件说明
- `app.py`：Streamlit 网页端主程序（内置 OO、SGH、RB、LC 四大品牌精准替换引擎 + 华东师范大学 `ecnu-plus` / `ecnu-max` 接口集成 + 本地正则兜底）。
- `requirements.txt`：云端自动安装的 Python 依赖列表。
- `templates/`：已整理好的四个品牌标准母版（`OO.pptx`, `SGH.pptx`, `RB.pptx`, `LC.pptx`）。

## 2. 零成本云端部署步骤（Streamlit Community Cloud，同事点链接就能用）
1. 注册并登录 [GitHub](https://github.com/)，新建一个仓库（Repository，建议设为 Private 私有仓库保护公司内部文件）。
2. 把本压缩包解压后的 **`app.py`**、**`requirements.txt`** 以及整个 **`templates` 文件夹** 上传到该 GitHub 仓库根目录。
3. 打开 [Streamlit Community Cloud (share.streamlit.io)](https://share.streamlit.io/)，用 GitHub 账号登录，点击 **Create app**。
4. 选择刚才创建的仓库，Main file path 填 `app.py`。
5. 点击 **Advanced settings...** -> 在 **Secrets** 框里填入你在华师大平台申请的 API 令牌（这样同事打开网页直接就能用，不用她们自己填 Token）：
   ```toml
   ECNU_API_KEY = "你的华师大API令牌"
   ```
6. 点击 **Deploy!**，等待约 1 分钟即可生成一个永久免费的专属网页链接，直接发给部门同事收藏即可使用。
