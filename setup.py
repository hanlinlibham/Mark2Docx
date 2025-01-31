from setuptools import setup, find_packages

setup(
    name="markdocx",
    version="0.1.0",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    install_requires=[
        "python-docx==0.8.11",
        "Markdown>=3.4.0",
        "beautifulsoup4==4.10.0",
        "PyYAML>=5.4.1",
        "requests>=2.25.1"
    ],
    package_data={
        'markdocx': ['config/*.yaml'],
    },
    python_requires='>=3.6',
    author="Your Name",
    author_email="your.email@example.com",
    description="Convert Markdown to Word documents with customizable styles",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/markdocx",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
) 