Game Boy Camera PicNRec interface for not-just-Windows.

See the [project page](https://nyuu.page/projects/picnrec/) for details.

# Requirements (other than those in the `toml`)

- USB-A to USB-C cable (This is *really* important, C-to-C no worky)

# Run

Pick one of:

* `uv run picnrec-gui` (*considerably* more featureful than the cli tool)

* `uv run picnrec`

Example cli commands:

```
uv run picnrec info
uv run picnrec export --all
uv run picnrec erase
```
