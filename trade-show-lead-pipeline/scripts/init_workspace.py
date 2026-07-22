"""一键初始化工作文件夹。用法: python init_workspace.py <文件夹路径>

创建文件夹结构、客户表模板、知识库占位文档、交谈要点采集模板。
已存在的文件不会被覆盖。
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = sys.argv[1] if len(sys.argv) > 1 else "."

DIRS = ["01-名片照片/待处理", "01-名片照片/已处理", "02-知识库/原始素材"]

COMPANY = """# 公司介绍（知识库）

> ⚠️ 占位文档。把公司真实资料放进「02-知识库/原始素材」，然后对Claude说"整理知识库"，
> 它会整理成正式版。整理前开发信只能用通用信息，效果打折。

## 公司概况
- 公司名称：
- 主营业务 / 核心产品线：
- 目标市场与客户类型：
- 成立年份 / 规模 / 工厂情况：

## 核心卖点（写开发信的弹药）
- （例：自有工厂，支持OEM/ODM）
- （例：已通过的认证、已服务的渠道类型）

## 红线（开发信里绝不能出现）
- 不透露其他客户的销售数据与合作细节
- 不承诺具体价格
- （补充你们公司的其他红线）
"""

TEMPLATES = """# 开发信范例与风格标准（知识库）

> ⚠️ 占位文档。把拿到过回复的优秀开发信放进「02-知识库/原始素材」，说"整理知识库"即可提炼成正式版。

## 风格标准
- 语言：英文，商务但自然，像一对一写的，不像群发
- 长度：100-160词，3-4个短段落
- 结构：①展会钩子（见面场景+对方感兴趣的产品）②我们是谁（只挑最相关的1个卖点）
  ③价值点/社会证明 ④轻CTA（发目录或约15分钟通话，二选一）
- 主题行：短、含对方公司名或展会名或具体产品

## 禁忌
- 垃圾邮件词：free, best price, 100%, cheapest
- 首封不报价、不塞多个附件
- 不写"We are the manufacturer of..."式自我中心开头，先说对方

## 范例
（整理知识库后由AI填入你们公司的真实优秀范例）
"""

NOTES = """交谈要点采集
====================
用法：按名片照片的序号，一行一条，记下现场聊了什么、对方对什么产品感兴趣。
AI会自动按序号填进客户表的「展会备注」列，并用于开发信个性化。
写得越具体，开发信越像人写的。

1：
2：
3：

（示例）
3：对柚木砧板感兴趣，有自己的Amazon店铺，要电商包装报价
"""

os.makedirs(ROOT, exist_ok=True)
for d in DIRS:
    os.makedirs(os.path.join(ROOT, d), exist_ok=True)

files = {
    "02-知识库/公司介绍.md": COMPANY,
    "02-知识库/开发信范例.md": TEMPLATES,
    "01-名片照片/待处理/交谈要点采集.txt": NOTES,
}
for rel, content in files.items():
    path = os.path.join(ROOT, rel)
    if os.path.exists(path):
        print(f"跳过（已存在）: {rel}")
        continue
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"已创建: {rel}")

table = os.path.join(ROOT, "客户表.xlsx")
if os.path.exists(table):
    print("跳过（已存在）: 客户表.xlsx")
else:
    subprocess.run([sys.executable, os.path.join(HERE, "build_table.py"), table], check=True)

print(f"\n✅ 工作文件夹已就绪: {ROOT}")
print("下一步：把公司资料放进 02-知识库/原始素材/，对Claude说「整理知识库」")
