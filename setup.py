from setuptools import setup


def readme():
    with open('README.md') as f:
        return f.read()

setup(
        name='msgraphapi',
        version='0.1',
        description='Python library interacting with the Microsoft Graph API',
        long_description=readme(),
        classifiers=[
            'Development Status :: 3 - Alpha',
            'Programming Language :: Python :: 2.7',
            'License :: Freely Distributable',
            'Natural Language :: English',
        ],
        url='https://github.com/jasonmlehman/msgraphapi.git',
        author='Jason Lehman',
        author_email='Jasonmlehman@yahoo.com',
        license='Freely Distributable',
        packages=[
            'msgraphapi',
            'msgraphapi.tools',
            'msgraphapi.creds'
            ],
        scripts=[
            'tools/listrolemembers'
        ],
        entry_points={
            'console_scripts': [
                'listrolemembers=msgraphapi.tools.listrolemembers:main'
            ],
        },
        install_requires=[
            'requests'
        ],
        zip_safe=False)