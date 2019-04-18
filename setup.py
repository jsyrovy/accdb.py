import setuptools

setuptools.setup(name="accdb",
                 version="1.0.1",
                 description="Microsoft Access Database Handler",
                 author="Jiri Syrovy",
                 author_email="jiri.syrovy@gmail.com",
                 packages=setuptools.find_packages(),
                 install_requires=["pypyodbc"],
                 zip_safe=False)
