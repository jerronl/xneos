from setuptools import setup, find_packages

setup(
    name='xneos',
    version='0.1.0',
    author='Jerron',
    author_email='jerron@gmail.com',
    description='A helping tool to submit jobs to NEOS from Excel',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/jerronl/xneos',  
    packages=find_packages(),
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
        'License :: OSI Approved :: MIT License',
    ],
    python_requires='>=3.7',
    install_requires=[
        'xlwings',
        'numpy'
    ],
    include_package_data=True, 
    package_data={
        "xneos": ["templates/*.xlsm", "templates/*.py"]
    },
    entry_points={
        'console_scripts': [
            'xneos = xneos.quickstart:main'
        ],
    },
)

