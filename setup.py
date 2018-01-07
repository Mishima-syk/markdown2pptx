from setuptools import setup

setup(
    name='md2pptx',
    version='0.1',
    py_modules=['md2pptx'],
    install_requires=[
        'Click',
    ],
    entry_points='''
        [console_scripts]
        md2pptx=md2pptx:cli
    ''',
)
