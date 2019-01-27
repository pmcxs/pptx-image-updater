# PPTX Image Updater

One Paragraph of project description goes here

## Getting Started

Setting up this command locally is very easy

### Prerequisites

This tool requires python installed (compatible with either 2.7+ or 3.3+).

For additional info please refer to https://wiki.python.org/moin/BeginnersGuide/Download

### Installing

- Start by cloning this repo to your computer

- Install [python-pptx](https://python-pptx.readthedocs.io/en/latest/user/install.html) module

```
pip install python-pptx
``` 
- Navigate to the newly created folder and run the following command to validate that it's working:

```
python update-images.py -h
```
This should display the various usage options of the tool

## How it works

The concept of this tool is very simple:

- you have a PPTX which includes images which you want to update from a certain source. For example, suppose you have an image that represents a calendar that gets updated quite often and you have to manually update it on the PPTX each time you're presenting it

- each image that you want to have automatically updated should include an hyperlink to the relevant source

- then just run the tool like:

```
python update-images.py <yourfile>.pptx
```

- it will automatically:
    - open the file 
    - iterate the various slides
    - for each image validate that it has an hyperlink
    - if so, download the image from that url and update it on the slide
    - save the file again

### Authentication ###

- Some images might be on the intranet under some sort of authentication. This tool also supports sending basic auth data by supplying the --username (-u) and --password (-p) parameters

```
python update-images.py <yourfile>.pptx -u myuser -p mypassword
```

If the password has special characters (like '!') it should be placed under single-quotes, like:

```
python update-images.py <yourfile>.pptx -u myuser -p 'mypass123!!!'
```

### Known Limitations ###
- Currently only basic-auth is supported. If you need another type of authentication it would need to be implemented

- If you duplicate the same image on the PPTX file and assign them different hyperlinks, it would update both images with the same data, particularly due to an optimization done by Powerpoint, on which it would store both images blob on the same place


## Contributing

No major guidelines on this. Just be reasonable :)

## Authors

* **Pedro Sousa** - *Initial work* - [pmcxs](https://github.com/PurpleBooth)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Hat tip to anyone whose code was used
* Inspiration
* etc
