#!/bin/bash
cd "$(dirname "$0")"

echo "============================================"
echo "  生产成本结转凭证生成工具"
echo "============================================"
echo ""
echo "当前目录下的 Excel 文件："
ls *.xlsx 2>/dev/null
echo ""
read -p "请输入文件名（如 生产成本0713.xlsx）：" INPUT_FILE

if [ -z "$INPUT_FILE" ]; then
    echo "未输入文件名，退出。"
    read -p "按回车键退出..."
    exit 1
fi

python3 generate_voucher.py "$INPUT_FILE"
echo ""
read -p "按回车键退出..."
