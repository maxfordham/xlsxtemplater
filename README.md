# xlsxtemplater

high-level wrapper that sits on xlsxwriter to template the output of the pandas dataframes to formatted excel tables.


## build 

### mount conda channel

- mount the network location

```bash
mkdir /mnt/conda-bld
sudo mount -t drvfs '\\barbados\apps\conda\conda-bld' /mnt/conda-bld
```

### conda build

```bash
wsl
conda activate base_mf
# add conda mf conda channel if not already there... check with command below...
# conda config --show
# run sh command from root dir of package
conda build conda.recipe

```

- note. builds here: `\\wsl$\Ubuntu-20.04\home\gunstonej\miniconda3\envs\base_mf\conda-bld` unless `--croot /mnt/c/engDev/channel` is specified
- once built check its working, then publish to MF network

### publish to MF

- copy and paste the linux-64 files `*.tar.bz2` into `\\barbados\apps\conda\conda-bld\linux-64`
- and convert to all platforms

```bash
conda convert --platform all /mnt/conda-bld/linux-64/xlsxtemplater-v0.1.5*.tar.bz2
conda index /mnt/conda-bld
```

- install to wsl from network channel

```bash
conda config --add channels file:///mnt/conda-bld
mamba install xlsxtemplater
# or 
mamba install -c file:///mnt/conda-bld xlsxtemplater
```

- install to windows from network channel

to install conda packages into windows you need to expose the conda channel.

```{cmd}
:: navigate to Z:\conda\conda-bld
:: create a python server of the directory on your local host
Z:\conda\conda-bld> python -m http.server
:: open a new cmd
```

```cmd
mamba install mypackage -c http://localhost:8000/
```
