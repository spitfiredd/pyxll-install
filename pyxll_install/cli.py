import os
import pathlib
import argparse
import logging

from enum import Enum

from pyxll_install.install import add_pyxll_registry_keys
from pyxll_install.uninstall import uninstall_all


logger = logging.getLogger('Pyxll Install Cli')
logger.setLevel(logging.DEBUG)


class InstallOptions(Enum):
    install = 'install'
    uninstall = 'uninstall'

    def __str__(self):
        return self.value


def cli():
    parser = argparse.ArgumentParser('Pyxll Installation')
    parser.add_argument('install', type=InstallOptions,
                        choices=list(InstallOptions))
    parser.add_argument('-i', '--input', dest='path', default=None,
                        help='File path of pyxll.xll file, default looks for \
                             a .pyxll folder in userprofile.')

    opts = parser.parse_args()
    if opts.install == InstallOptions.install:
        if opts.path:
            path = pathlib.Path(opts.path)
            ext = path.suffix
            if ext != '.xll':
                logger.error(f'Please provide a valid pyxll.xll file.')
                logger.error(f'Supplied Path: {path}')
                logger.error(f'Invalid Extension: {ext}')
            else:
                logger.info(f'Adding Registry keys {path}')
                add_pyxll_registry_keys(path)
        else:
            logger.info('No File path provided, using default location.')
            base_pyxll_path = pathlib.Path(os.getenv("USERPROFILE"))
            pyxll_path = base_pyxll_path / '.pyxll' / 'pyxll.xll'
            assert os.path.exists(pyxll_path)
            logger.info(f'Default Path: {pyxll_path}')
            add_pyxll_registry_keys(pyxll_path)
    elif opts.install == InstallOptions.uninstall:
        logger.info('Removing Pyxll Registry Keys')
        uninstall_all()


if __name__ == "__main__":
    cli()
