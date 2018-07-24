function checkForPuTTy()

a = system('C:\Program Files\PuTTY\plink.exe')
if a == 1
  oFN = websave('putty.msi','https://the.earth.li/~sgtatham/putty/latest/w64/putty-64bit-0.70-installer.msi');
  system(oFN);
  system(['del ' oFN]);
end
