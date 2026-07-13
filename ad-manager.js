/**
 * Advanced AdSense Manager
 * Dynamically fetches ad placeholders and injects them into optimal layout slots 
 * across the entire ilovespdfs.in platform to maximize AdSense revenue.
 */

document.addEventListener("DOMContentLoaded", function () {
    const basePath = window.location.pathname.startsWith('/seo_pages/') ? '../ads/' : '/ads/';

    async function loadAdSnippet(filename) {
        try {
            const response = await fetch(basePath + filename);
            if (response.ok) {
                return await response.text();
            }
        } catch (e) {
            console.warn("Could not load ad snippet: " + filename);
        }
        return null;
    }

    function insertAd(html, targetSelector, position = 'afterend', fallbackSelector = null) {
        if (!html) return;
        
        // Wrap the ad in a responsive container to prevent Layout Shifts (CLS)
        const adContainer = document.createElement('div');
        adContainer.className = 'dynamic-ad-container';
        adContainer.style.margin = '20px auto';
        adContainer.style.maxWidth = '1000px';
        adContainer.style.textAlign = 'center';
        adContainer.style.overflow = 'hidden';
        adContainer.style.minHeight = '90px'; // Prevent CLS for standard banner
        adContainer.innerHTML = html;

        let target = document.querySelector(targetSelector);
        if (!target && fallbackSelector) {
            target = document.querySelector(fallbackSelector);
        }

        if (target) {
            if (position === 'afterend') {
                target.parentNode.insertBefore(adContainer, target.nextSibling);
            } else if (position === 'beforebegin') {
                target.parentNode.insertBefore(adContainer, target);
            } else if (position === 'beforeend') {
                target.appendChild(adContainer);
            }
        }
    }

    async function initializeAds() {
        // 1. Load Auto Ads (Header)
        const autoAds = await loadAdSnippet('auto-ads.html');
        if (autoAds) {
            const headContainer = document.createElement('div');
            headContainer.innerHTML = autoAds;
            // Execute any scripts found in auto-ads
            Array.from(headContainer.querySelectorAll('script')).forEach(oldScript => {
                const newScript = document.createElement('script');
                Array.from(oldScript.attributes).forEach(attr => newScript.setAttribute(attr.name, attr.value));
                newScript.appendChild(document.createTextNode(oldScript.innerHTML));
                document.head.appendChild(newScript);
            });
        }

        // 2. Load Top Banner (High Visibility, Above the Fold)
        const topBanner = await loadAdSnippet('top-banner.html');
        // Insert below breadcrumb or below the main header
        insertAd(topBanner, 'nav[aria-label="breadcrumb"]', 'afterend', 'header');

        // 3. Load Bottom Banner (Below workspace or grid)
        const bottomBanner = await loadAdSnippet('bottom-banner.html');
        insertAd(bottomBanner, '.tool-workspace', 'afterend', '.tools-grid');

        // 4. Load In-Article Ad (Inside SEO content)
        const inArticle = await loadAdSnippet('in-article.html');
        const seoContent = document.querySelector('.seo-content');
        if (seoContent && inArticle) {
            const paragraphs = seoContent.querySelectorAll('p');
            if (paragraphs.length >= 2) {
                // Insert after the first paragraph
                insertAd(inArticle, 'p:first-of-type', 'afterend');
            }
        }
        
        // Execute scripts inside injected ad containers (AdSense requires push{})
        setTimeout(() => {
            document.querySelectorAll('.dynamic-ad-container script').forEach(oldScript => {
                if(oldScript.innerHTML.includes('adsbygoogle')) {
                    try {
                        eval(oldScript.innerHTML);
                    } catch(e) {}
                }
            });
        }, 500);
    }

    initializeAds();
});
