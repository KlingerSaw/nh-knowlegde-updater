/**
 * HTML cleaning utility functions
 */

import { decode } from 'html-entities';

export class HtmlCleaner {
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
    "&#160;": " "
  };

  /**
   * Clean HTML text by removing tags and converting entities
   * @param html The HTML string to clean
   * @returns Cleaned plain text
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

    // Decode any remaining HTML entities
    cleaned = decode(cleaned);
    
    return cleaned.trim();
  }

  /**
   * Clean an object's string properties recursively
   * @param obj The object to clean
   * @returns Object with cleaned string values
   */
  static cleanObject(obj: any): any {
    if (obj === null || obj === undefined) {
      return obj;
    }

    if (typeof obj === 'string') {
      return this.cleanHtmlText(obj);
    }

    if (Array.isArray(obj)) {
      return obj.map(item => this.cleanObject(item));
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