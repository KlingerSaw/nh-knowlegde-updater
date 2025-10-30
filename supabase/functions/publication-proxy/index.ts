const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
};

/**
 * HTML cleaning utility functions
 */
class HtmlCleaner {
  private static readonly HTML_ENTITIES: { [key: string]: string } = {
    // Non-breaking space
    "\u00a0": " ",
    
    // Danish characters - lowercase
    "&aelig;": "æ",
    "&#230;": "æ",
    "&oslash;": "ø", 
    "&#248;": "ø",
    "&aring;": "å",
    "&#229;": "å",
    
    // Danish characters - uppercase
    "&AElig;": "Æ",
    "&#198;": "Æ",
    "&Oslash;": "Ø",
    "&#216;": "Ø",
    "&Aring;": "Å",
    "&#197;": "Å",
    
    // Common HTML entities
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": '"',
    "&apos;": "'",
    "&#39;": "'",
    "&nbsp;": " ",
    "&#160;": " ",
    
    // Additional common entities
    "&ndash;": "–",
    "&mdash;": "—",
    "&hellip;": "…",
    "&laquo;": "«",
    "&raquo;": "»",
    "&ldquo;": "\"",
    "&rdquo;": "\"",
    "&lsquo;": "'",
    "&rsquo;": "'",
  };

  /**
   * Clean HTML text by removing tags and converting entities
   */
  static cleanHtmlText(html: string): string {
    if (!html || typeof html !== 'string') {
      return '';
    }

    let cleaned = html;

    // Replace HTML entities with proper characters
    for (const [entity, replacement] of Object.entries(this.HTML_ENTITIES)) {
      cleaned = cleaned.replace(new RegExp(entity, 'g'), replacement);
    }

    // Remove link tags but keep the text content
    cleaned = cleaned.replace(/<a [^>]*>(.*?)<\/a>/gi, '$1');
    
    // Remove all other HTML tags
    cleaned = cleaned.replace(/<[^>]+>/g, '');
    
    // Replace carriage returns and newlines with spaces
    cleaned = cleaned.replace(/\r/g, '').replace(/\n/g, ' ');
    
    // Collapse multiple whitespace characters into single spaces
    cleaned = cleaned.replace(/\s+/g, ' ');
    
    return cleaned.trim();
  }

  /**
   * Clean an object's string properties recursively
   */
  static cleanObject(obj: any): any {
    if (obj === null || obj === undefined) {
      return obj;
    }

    if (typeof obj === 'string') {
      return this.cleanHtmlText(obj);
    }

    if (Array.isArray(obj)) {
      return obj.map((item: any) => this.cleanObject(item));
    }

    if (typeof obj === 'object') {
      const cleaned: any = {};
      for (const [key, value] of Object.entries(obj)) {
        cleaned[key] = this.cleanObject(value);
      }
      return cleaned;
    }

    return obj;
  }
}

/**
 * Validate if a URL is safe to fetch (not local or private network)
 */
function validateUrl(urlString: string): { isValid: boolean; error?: string } {
  try {
    const url = new URL(urlString);
    const hostname = url.hostname.toLowerCase();

    // Check for localhost variations
    if (hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '::1') {
      return { isValid: false, error: 'Access to localhost is not allowed' };
    }

    // Check for private IP ranges
    const ipv4Regex = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/;
    const ipv4Match = hostname.match(ipv4Regex);
    
    if (ipv4Match) {
      const [, a, b, c, d] = ipv4Match.map(Number);
      
      // Check private IP ranges
      if (
        (a === 10) || // 10.0.0.0/8
        (a === 172 && b >= 16 && b <= 31) || // 172.16.0.0/12
        (a === 192 && b === 168) || // 192.168.0.0/16
        (a === 169 && b === 254) || // 169.254.0.0/16 (link-local)
        (a === 127) // 127.0.0.0/8 (loopback)
      ) {
        return { isValid: false, error: 'Access to private IP addresses is not allowed' };
      }
    }

    // Only allow HTTP and HTTPS protocols
    if (url.protocol !== 'http:' && url.protocol !== 'https:') {
      return { isValid: false, error: 'Only HTTP and HTTPS protocols are allowed' };
    }

    return { isValid: true };
  } catch (error) {
    return { isValid: false, error: 'Invalid URL format' };
  }
}

/**
 * Fetch with retry logic for better reliability
 */
async function fetchWithRetry(
  url: string, 
  options: RequestInit = {}, 
  maxRetries: number = 3,
  baseDelay: number = 1000
): Promise<Response> {
  let lastError: Error;
  
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      console.log(`Fetch attempt ${attempt + 1}/${maxRetries + 1} for: ${url}`);
      
      const response = await fetch(url, {
        ...options,
        signal: AbortSignal.timeout(30000), // 30 second timeout
      });
      
      // Return successful responses immediately
      if (response.ok) {
        return response;
      }
      
      // For non-2xx responses, throw an error to trigger retry logic
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));
      console.error(`Fetch attempt ${attempt + 1} failed:`, lastError.message);
      
      // Don't retry on the last attempt
      if (attempt === maxRetries) {
        break;
      }
      
      // Calculate delay with exponential backoff
      const delay = baseDelay * Math.pow(2, attempt);
      console.log(`Retrying in ${delay}ms...`);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
  
  throw lastError;
}

Deno.serve(async (req: Request) => {
  // Handle CORS preflight requests
  if (req.method === "OPTIONS") {
    return new Response(null, {
      status: 204,
      headers: corsHeaders,
    });
  }

  try {
    // Get the site URL and GUID from the query parameters
    const url = new URL(req.url);
    const siteUrl = url.searchParams.get("url");
    const guid = url.searchParams.get("guid");

    if (!siteUrl || !guid) {
      return new Response(
        JSON.stringify({ error: "Missing site URL or GUID parameter" }),
        {
          status: 400,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        }
      );
    }

    // Validate the site URL
    const validation = validateUrl(siteUrl);
    if (!validation.isValid) {
      return new Response(
        JSON.stringify({ error: validation.error }),
        {
          status: 403,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        }
      );
    }

    // Construct the publication API URL
    const publicationUrl = `${siteUrl.replace(/\/$/, "")}/api/publication/${guid}`;
    console.log(`Fetching publication from: ${publicationUrl}`);

    // Validate the constructed URL as well
    const publicationValidation = validateUrl(publicationUrl);
    if (!publicationValidation.isValid) {
      return new Response(
        JSON.stringify({ error: publicationValidation.error }),
        {
          status: 403,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        }
      );
    }

    // Fetch the publication
    const response = await fetchWithRetry(publicationUrl, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
      },
    }, 3, 1000);

    // Get the JSON content
    console.log("Parsing JSON response...");
    const data = await response.json();
    console.log("JSON parsed successfully");

    // Clean HTML from the publication data and return
    const cleanedData = HtmlCleaner.cleanObject(data);
    
    return new Response(
      JSON.stringify(cleanedData),
      {
        status: 200,
        headers: {
          ...corsHeaders,
          "Content-Type": "application/json",
        },
      }
    );
  } catch (error) {
    console.error("Error processing request:", error);
    return new Response(
      JSON.stringify({ error: error?.message || String(error) || "Unknown error occurred" }),
      {
        status: 500,
        headers: {
          ...corsHeaders,
          "Content-Type": "application/json",
        },
      }
    );
  }
});