# WINSTON 🎩

**Work Information Search Tool for Obscure topics & Niches**

Finds postdoctoral and early-career research posts that suit **you**, and tells you whether you are eligible for them.

It searches job sites and vacancy pages, ranks every ad against your CV and a short description of what you want, then has a language model **running on your own computer** read the best ads in full. For each ad the model rates the field fit and extracts the PhD rule ("PhD awarded by the start date", "thesis submitted by 18 February"…). The script then checks that rule against your PhD timeline. The result is a spreadsheet, best matches first.

Nothing about you is sent anywhere: your CV and profile are only read by models on your machine.

## How it works

1. **Collect.** Runs keyword searches on job sites (jobs.ac.uk, EURAXESS…) and reads vacancy listing pages (Cambridge colleges, Oxford Gazette, H-Net…). Ads are cached, so nothing is downloaded twice.
2. **Rank.** Turns your profile, your CV and every ad into embeddings and scores how close each ad is to you. This is a fast first sieve.
3. **Judge.** A local language model reads the top ads in full. It rates field fit (0–10), reads the career level and start date, and quotes the PhD rule with its dates. The script, not the model, compares those dates with your timeline and decides `yes` / `no` / `check` / `unclear`.

All the output goes into `postdoc_matches.xlsx`.

## Installation

You need **Python 3.11 or newer**.

```bash
git clone https://github.com/jmgr2553/jobHunter.git
cd jobHunter
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

For step 3 you also need a local model server. The simplest is [Ollama](https://ollama.com):

```bash
ollama pull qwen3.5:4b
```

You can also use LM Studio, llama.cpp or any other server with an OpenAI-style API (see [The language model](#the-language-model)). Step 3 is optional: `--no-llm` skips it.

## Try it with the example

The repository includes a fictional candidate, Dr John Doe, a modern historian. His files are in `examples/john_doe/`.

```bash
cp config.example.toml config.toml
python hunter.py
```

The first run downloads the embedding model (~1 GB) and can take a while. Later runs only download new ads.

## Set it up for yourself

1. Copy the example files to a private folder:
   ```bash
   cp -r examples/john_doe my_data
   ```
2. Replace or edit the files in `my_data/`:
   - **`cv.pdf`**: your CV (a `.pdf` or a plain `.txt` file).
   - **`profile.txt`**: 100–300 words on the posts you want, using the words ads would use.
   - **`candidate.txt`**: facts that decide eligibility: PhD status and dates, earliest start date, nationality and right to work, languages, degrees. The model reads this.
   - **`not_interested.txt`** (optional): posts you do not want. Ads that resemble it are pushed down the ranking.
3. In `config.toml`:
   - point `[files]` at `my_data/`;
   - set `[timeline]` to your thesis submission and PhD award dates;
   - describe your fields in `[fit_scale]`;
   - set the search `queries` for your subject.

`config.toml` and `my_data/` are listed in `.gitignore`, so they stay out of git.

## Usage

```bash
python hunter.py                    # all three steps
python hunter.py --no-llm           # skip step 3 (no language model needed)
python hunter.py --rejudge          # re-read every ad (e.g. after editing candidate.txt)
python hunter.py --model llama3.1:8b   # use another model for this run
python hunter.py --config other.toml   # use another settings file
python hunter.py --check-sources    # test every website once, then stop
```

## Choosing the websites

The sources are listed in `config.toml`. There are two kinds.

**Search sites** have a keyword search. You give the address of a results page, with `{query}` where the keywords go and `{page}` where the page number goes, plus a pattern for what an ad's address looks like:

```toml
[[search_sites]]
name = "jobs.ac.uk"
url = "https://www.jobs.ac.uk/search/?keywords={query}&pageSize=25&startIndex={page}"
first_page = 1          # the first page's number
page_step = 25          # how the number grows from page to page
per_page = 25
link_pattern = 'https?://[^/]+/job/[A-Za-z0-9]+/[^/?#]*'
queries = ["history", '"research fellowship"']
```

**Listing pages** are single pages that list current vacancies. The script keeps every link whose text looks like a postdoc or fellowship (see `[titles]`):

```toml
[[listing_pages]]
name = "Oxford Gazette"
url = "https://www.ox.ac.uk/about/how-we-are-run/gazette/jobs"
always_read = true      # the model reads all its ads, however they rank
```

To add a site:

1. Copy one of the entries and change it.
2. Run `python hunter.py --check-sources`. It loads each source once and shows what it finds.

Add `enabled = false` to switch a source off without deleting it. Every option is explained in `config.example.toml`.

**Limitation:** some sites build their job list with JavaScript after the page loads. This script can't read those, and `--check-sources` warns you about them. Often the same posts also appear on a site it can read.

## The language model

Configure it in the `[llm]` section of `config.toml`.

**Ollama** (default):

```toml
[llm]
provider = "ollama"
base_url = "http://localhost:11434"
model = "qwen3.5:4b"
```

**LM Studio, llama.cpp, vLLM, Jan, LocalAI**, or any server with an OpenAI-style API:

```toml
[llm]
provider = "openai"
base_url = "http://localhost:1234/v1"   # LM Studio's default; llama-server uses :8080/v1
model = "qwen3-8b"                      # the name the server gives the model
```

How to choose a model:

- **Size.** Models of 4–8 billion parameters run on a laptop with 16 GB of memory. Each ad takes 25–60 seconds. Larger models judge better but read more slowly.
- **Instruction-following.** Pick a model that follows instructions well, such as Qwen or Llama 3.1 and later.
- **Changing models.** If you change the model, the ads are read again with the new one. Answers are stored in the cache, so an interrupted run (Ctrl+C) loses nothing.

## The spreadsheet

| Column | Meaning |
|---|---|
| Overall | The model's field fit (0–10), lowered for posts you can't take or at the wrong career level |
| Eligible | `yes`, `no`, `check` (another requirement may rule you out) or `unclear` (no PhD rule found) |
| Eligibility note | Why, with the ad's PhD rule quoted |
| Fit score | The step-2 embedding score |
| Deadline, Days left, Link, Also at | The same post found on several sites is merged into one row |

**Always check an ad yourself before applying.** The model can misread an advert, and some ads leave out the funder's rules. You can add those rules under `[[scheme_notes]]`.

## Please be polite to the websites

The script waits between downloads (`[network] seconds_between_downloads`) and backs off when a site asks it to. Keep the delays, don't run it in a tight loop, and respect each site's terms of use. Once a day or once a week is plenty for job hunting.

## Licence

MIT - see [LICENSE](LICENSE).
