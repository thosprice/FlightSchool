Repository contains the following files:<br>
 X-Plane 9      - folder containing 29 files to set preferences, update airport map and naviads, and load situations into X-Plane 9<br>
 manifest.txt   - List of path\file for each X-Plane 9 file<br>
 hashes.txt     - File of SHA256 hashes generated, corresponding to each file in the manifest.<br>
 hash_all.bat   - Run this from the directory containing the X-Plane 9 files in VirtualStore to generate Windows-compatible SHA256 hashes for the repo<br>
 Potential_Issues - list of possible admin changes that might cause download to fail in the future, AI generated while developing the repo/install script.<br>
 .gitattributes - set so that files in the manifest are treated as binary. They can't be edited directly in the repo, but this keeps github from doing things 
  like swapping <cr> for <cr><lf> pairs, that would break the download validation script.
