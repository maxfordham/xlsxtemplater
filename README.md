# xlsxtemplater

high-level wrapper that sits on xlsxwriter to template the output of the pandas dataframes to formatted excel tables.

## build and publish to MF using conda-build

- build locally using wsl

```bash
wsl
conda activate base_mf
conda build conda.recipe
conda build conda.recipe --croot /mnt/c/engDev/channel
```

- once built check its working, then publish to MF network
- mount the network location

```bash
mkdir /mnt/barbados
sudo mount -t drvfs '\\barbados\apps\conda\conda-bld' /mnt/conda-bld
```

- copy and paste the linux-64 files `*.tar.bz2` into `\\barbados\apps\conda\conda-bld\linux-64`
- and convert to all platforms

```bash
conda convert --platform all /mnt/conda-bld/linux-64/xlsxtemplater*.tar.bz2
conda index /mnt/conda-bld
```

- install from network channel

```bash
conda install -c file:///mnt/conda-bld xlsxtemplater
```