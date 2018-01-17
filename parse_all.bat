rem Convert all the starship files to JSON. ~dpn translates to drive,
rem path, base filename. Quoting the paths is necessary because the filenames
rem contain spaces.
for %%f in (ships\*.swc) do (
	python parse_starship.py "%%f" > "%%~dpnf.json"
)
