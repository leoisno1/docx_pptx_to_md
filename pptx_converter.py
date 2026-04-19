import logging
import os
from pathlib import Path

from entry import convert
from log import setup_logging
from custom_types import ConversionConfig

setup_logging(compat_tqdm=True)
logger = logging.getLogger(__name__)


def pptx_to_markdown(pptx_path, output_path, image_dir=None):
    """将PPTX文件转换为Markdown格式

    Args:
        pptx_path: PPTX文件路径
        output_path: 输出MD文件路径
        image_dir: 图片输出目录，默认为output_path同目录下的img文件夹
    """
    try:
        if image_dir is None:
            image_dir = Path(output_path).parent / 'img'

        config = ConversionConfig(
            pptx_path=Path(pptx_path),
            output_path=Path(output_path),
            image_dir=Path(image_dir),
            disable_image=False,
            disable_wmf=True,  # Linux环境下禁用wmf转换
            disable_color=True,  # 不添加颜色标签
            disable_escaping=False,
            disable_notes=True,  # 不添加备注
            enable_slides=False,
            is_wiki=False,
            min_block_size=15,
        )

        convert(config)
        return True
    except Exception as e:
        logger.error(f"PPTX转换失败: {e}")
        return False
