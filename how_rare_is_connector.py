import logging
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timezone
import pytz


class HowRareIs:
    _URL = "https://howrare.is/drops"

    def __init__(self):
        self._log = logging.getLogger(__name__)

    def _get_page_html(self):
        r = requests.get(self._URL)
        self._log.info("Downloading drops...")
        self._log.debug("Retreiving content from: %s", self._URL)

        if r.status_code != 200:
            raise RuntimeError(
                f"Unable to retrieve information from {self._URL} (Status: {r.status_code})"
            )

        content = r.content
        self._log.info("Done.")
        return content

    def _get_soup(self, page_text: str):
        return BeautifulSoup(page_text, "html5lib")

    def _validate_utc(self, utc_str):
        if "utc" not in utc_str.lower():
            self._log.debug(
                "Invalid UTC time string encountered: %s. ",
                utc_str,
            )

        try:
            utc_time = datetime.strptime(utc_str.lower(), "%H:%M utc")
            return utc_time.strftime("%I:%M %p")
        except ValueError as e:
            self._log.debug("Unable to parse UTC time string. %s", repr(e))
            return None

    def _utc_str_to_est(self, utc_str):
        from_zone = timezone.utc
        to_zone = pytz.timezone("America/New_York")

        utc = None
        try:
            utc = datetime.strptime(utc_str.lower(), "%H:%M utc")
        except ValueError as e:
            self._log.debug("Unable to parse UTC time string. %s", repr(e))
            return None

        utc = utc.replace(tzinfo=from_zone, year=datetime.now().year)
        utc = utc.astimezone(to_zone)

        return utc.strftime("%I:%M %p")

    def get_drops(self):
        """Retrieve all upcoming Solana NFT drops.

        @return A dict of drops, indexed by date.
        @example
        ```python
            {
                "JANUARY 25TH":
                [
                    {
                        "project_name": str,
                        "time_est": str,
                        "time_utc": str,
                        "twitter_url": str,
                        "discord_url": str,
                        "website_url": str,
                        "supply": int,
                        "mint_price": float
                    },
                    ...
                ],
                ...
            }
        ```
        """
        drops = {}
        soup = self._get_soup(self._get_page_html())

        drop_elems = soup.find_all("div", class_="all_collections")
        self._log.debug("Found %i drops.", len(drop_elems))

        log_count = 0
        drop_count = 0

        for element in drop_elems:
            date = element.find_all("div", class_="drop_date")
            if len(date) > 0:
                date = date[0].text.strip()
                try:
                    format_date = datetime.strptime(date[:-2], "%B %d")
                    date = format_date.strftime("%m/%d")
                except ValueError as e:
                    self._log.debug("Invalid date format: %s", repr(e))
            else:
                self._log.warning(
                    "Unable to parse HTML to find a date. "
                    "Will continue in hopes that this issue is only found on part of the page."
                )

            # Create a new record in the drops list
            drops[date] = []

            # Now, we'll have to go through each drop
            info_elements = element.find_all("div", class_="all_coll_row")
            for drop in info_elements:
                log_count += 1
                log_count = log_count % 3

                logging.StreamHandler.terminator = "\r"
                self._log.info("{}".format("." * log_count).ljust(10, " "))
                logging.StreamHandler.terminator = "\n"

                drop_info = {
                    "project_name": None,
                    "time_est": None,
                    "time_utc": None,
                    "twitter_url": None,
                    "discord_url": None,
                    "website_url": None,
                    "supply": None,
                    "mint_price": None,
                }

                # Don't count the header row
                if any(name in ["legend", "drop_date"] for name in drop["class"]):
                    continue

                # This is a list of the information elements
                information = drop.find_all("div", class_="all_coll_col")
                has_time_till_mint = True if len(information) > 6 else False

                if len(information) > 5:
                    # Get Project Name
                    project_name = information[0]
                    project_name = project_name.find_all("span")
                    if len(project_name) > 0:
                        drop_info["project_name"] = project_name[0].text.strip()
                        self._log.debug(
                            "Project name found: %s", drop_info["project_name"]
                        )
                    else:
                        self._log.warning(
                            "Unable to find the project name for a drop. "
                            "Continuing anyway in hopes that this is not a problem."
                        )

                    # Get Project Links
                    project_links = information[1]
                    links = project_links.find_all("a", href=True)
                    urls = []
                    for link in links:
                        urls.append(link["href"].lower())
                    for url in urls:
                        if "twitter" in url:
                            self._log.debug("Found twitter_url: %s", url)
                            drop_info["twitter_url"] = url
                        elif "discord" in url:
                            self._log.debug("Found discord_url: %s", url)
                            drop_info["discord_url"] = url
                        else:
                            self._log.debug("Found website_url: %s", url)
                            drop_info["website_url"] = url

                    # Get Project Times
                    project_time = information[2].text.strip()
                    drop_info["time_est"] = self._utc_str_to_est(project_time)
                    drop_info["time_utc"] = self._validate_utc(project_time)
                    self._log.debug("Project_time (EST): %s", drop_info["time_est"])
                    self._log.debug("Project_time (UTC): %s", drop_info["time_utc"])

                    # Get Project Supply
                    supply = "Unknown"
                    try:
                        index = 4 if has_time_till_mint else 3
                        supply = int(information[index].text.strip())
                    except ValueError as e:
                        self._log.debug(
                            "Non-number supply value: %s",
                            repr(e),
                        )

                    drop_info["supply"] = supply
                    self._log.debug("Supply: %s", drop_info["supply"])

                    # Get Mint Price
                    index = 5 if has_time_till_mint else 4
                    mint_price = (
                        information[index]
                        .text.strip()
                        .lower()
                        .replace("sol", "")
                        .strip()
                    )
                    drop_info["mint_price"] = mint_price
                    self._log.debug("Mint Price: %s", drop_info["mint_price"])

                else:
                    self._log.warning(
                        "Unable to find information for a drop: Malformed drop div. "
                        "Continuing anyway in hopes that this is not a problem."
                    )

                drops[date].append(drop_info)
                drop_count += 1

        self._log.info("Found %s drops.", drop_count)
        return drops


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    con = HowRareIs()

    con.get_drops()
