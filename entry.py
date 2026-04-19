# Copyright 2024 Liu Siyao
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import logging

from outputter import MarkdownFormatter, WikiFormatter
from parser import parse
from custom_types import ConversionConfig
from utils import load_pptx, prepare_titles

logger = logging.getLogger(__name__)


def convert(config: ConversionConfig):
    if config.title_path:
        config.custom_titles = prepare_titles(config.title_path)

    prs = load_pptx(str(config.pptx_path))

    logger.info("conversion started")

    ast = parse(config, prs)

    if str(config.output_path).endswith('.json'):
        with open(config.output_path, 'w') as f:
            f.write(ast.model_dump_json(indent=2))
        logger.info(f'presentation data saved to {config.output_path}')
        return

    if config.is_wiki:
        out = WikiFormatter(config)
    else:
        out = MarkdownFormatter(config)

    out.output(ast)
    logger.info(f'converted document saved to {config.output_path}')
