import pyperclip
from googlesearch import search
import time

def google_and_copy_link(query):
    try:
        # Construct the query to explicitly search for the main domain
        query_string = f"{query} parfumo.net"
        # Perform the Google search
        search_results = search(query_string, num_results=1)
        # Extract the links from search results
        links = list(search_results)
        return links
    except Exception as e:
        print("An error occurred:", e)
        return []

def main():
    with open("links.txt", "w") as f:
        while True:
            query = input("Enter your search query (type 'quit' to exit): ")
            if query.lower() == 'quit':
                break
            
            queries = query.split(',')
            for q in queries:
                q = q.strip()  # Remove leading/trailing whitespaces
                links = google_and_copy_link(q)
                if links:
                    f.write(' '.join(links) + '\n')  # Write links with space separator
                    print(f"Search results for '{q}' saved to links.txt")
            
            # Introduce a delay of 2 seconds between searches
            time.sleep(2)

    print("Exiting...")

if __name__ == "__main__":
    main()
