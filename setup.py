from setuptools import setup, find_packages

setup(
    name='e2m_package',
    version='1.0',
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'e2m = e2m_package.main:main',
        ],
    },
    package_data={
        'e2m_package': ['templ/*', 'META-INF/*'],
    },
    author='huangleilei',
    author_email='huangll@sugon.com',
    description='禅道excel格式测试用例与xmind格式测试用例相互转换',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/GavenHwang/e2m_package.git',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.9',
    requires=['xmind', 'pandas'],
)
