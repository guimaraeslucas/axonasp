/*
 Copyright (c) 2003-2026, CKSource - Frederico Knabben. All rights reserved.
*/
CKEDITOR.addTemplates("default", {
    imagesPath: CKEDITOR.getUrl(CKEDITOR.plugins.getPath("templates") + "templates/images/"),
    templates: [
	{
		title: "Hero Section: Image Background",
		image: "hero.jpg",
		description: "A full-width hero banner with a background image and centered text.",
		html: 
			'<div style="position: relative; background-color: #222; background-image: url(\'https://quickersite.com/placeholder.jpg\'); background-size: cover; background-position: center; padding: 120px 20px; text-align: center; border-radius: 12px; margin-bottom: 30px; overflow: hidden; font-family: sans-serif;">' +
				'' +
				'<div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0, 0, 0, 0.5); z-index: 1;"></div>' +
				
				'' +
				'<div style="position: relative; z-index: 2; max-width: 800px; margin: 0 auto;">' +
					'<h1 style="color: #ffffff; font-size: 48px; font-weight: bold; margin: 0 0 20px 0; letter-spacing: -1px; line-height: 1.1;">' +
						'Create Your Digital Legacy' +
					'</h1>' +
					'<p style="color: #ffffff; font-size: 20px; line-height: 1.6; margin: 0 0 35px 0; opacity: 0.9; font-weight: 300;">' +
						'We provide the tools and expertise to turn your boldest ideas into reality. High-performance solutions for the modern era.' +
					'</p>' +
					'<div>' +
						'<a href="#" style="color: #ffffff; text-decoration: none; font-size: 18px; font-weight: bold; border: 2px solid #ffffff; padding: 12px 30px; border-radius: 4px; display: inline-block; transition: background 0.3s;">' +
							'Discover our process &rarr;' +
						'</a>' +
					'</div>' +
				'</div>' +
			'</div><p> </p>'
	},

        // 1. TWO COLUMN CARD GRID
        {
            title: "Two Column Cards",
            image: "2col.jpg",
            description: "Two responsive columns with header images and footer links.",
            html: 
                '<div style="display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 30px;">' +
                    '<div style="flex: 1; min-width: 280px; border: 1px solid #eee; border-radius: 8px; overflow: hidden; display: flex; flex-direction: column;">' +
                        '<img src="https://quickersite.com/placeholder.jpg" style="width:100%; height:150px; object-fit:cover;" />' +
                        '<div style="padding: 20px; flex-grow: 1;">' +
                            '<h4 style="margin-top:0;">Feature One</h4>' +
                            '<p style="color: #666;">Detailed description of your first service or feature goes here.</p>' +
                        '</div>' +
                        '<div style="padding: 15px; border-top: 1px solid #eee;">' +
                            '<a href="#" style="color: #0d6efd; font-weight: bold; text-decoration: none;">Learn more &rarr;</a>' +
                        '</div>' +
                    '</div>' +
                    '<div style="flex: 1; min-width: 280px; border: 1px solid #eee; border-radius: 8px; overflow: hidden; display: flex; flex-direction: column;">' +
                        '<img src="https://quickersite.com/placeholder.jpg" style="width:100%; height:150px; object-fit:cover;" />' +
                        '<div style="padding: 20px; flex-grow: 1;">' +
                            '<h4 style="margin-top:0;">Feature Two</h4>' +
                            '<p style="color: #666;">Detailed description of your second service or feature goes here.</p>' +
                        '</div>' +
                        '<div style="padding: 15px; border-top: 1px solid #eee;">' +
                            '<a href="#" style="color: #0d6efd; font-weight: bold; text-decoration: none;">Explore details &rarr;</a>' +
                        '</div>' +
                    '</div>' +
                '</div><p> </p>'
        },

        // 2. THREE COLUMN CARD GRID
        {
            title: "Three Column Cards",
            image: "3col.jpg",
            description: "Three responsive columns with header images and footer links.",
            html: 
                '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px;">' +
                    [1,2,3].map(i => 
                        '<div style="border: 1px solid #eee; border-radius: 8px; overflow: hidden; display: flex; flex-direction: column;">' +
                            '<img src="https://quickersite.com/placeholder.jpg" style="width:100%; height:120px; object-fit:cover;" />' +
                            '<div style="padding: 15px; flex-grow: 1;">' +
                                '<h4 style="margin: 0 0 10px 0;">Service ' + i + '</h4>' +
                                '<p style="color: #666; margin:0;">Brief summary of what this specific service offers.</p>' +
                            '</div>' +
                            '<div style="padding: 12px; background: #f9f9f9;">' +
                                '<a href="#" style="color: #0d6efd; font-weight: 600; text-decoration: none;">View Details</a>' +
                            '</div>' +
                        '</div>'
                    ).join('') +
                '</div><p> </p>'
        },

        // 3. IMAGE LEFT, TEXT RIGHT
        {
            title: "Split: Image Left",
            image: "iL.jpg",
            description: "Large image on the left, descriptive text on the right.",
            html: 
                '<div style="display: flex; flex-wrap: wrap; align-items: center; gap: 40px; margin-bottom: 40px;">' +
                    '<div style="flex: 1; min-width: 250px;">' +
                        '<img src="https://quickersite.com/placeholder.jpg" style="width:100%; border-radius: 12px;" />' +
                    '</div>' +
                    '<div style="flex: 1; min-width: 250px;">' +
                        '<h2 style="font-size: 28px; margin-bottom: 15px;">Focus on Growth</h2>' +
                        '<p style="font-size: 16px; line-height: 1.6; color: #555;">Our tools help you identify trends before they happen. Scale your business with data-driven confidence.</p>' +
                        '<a href="#" style="color: #0d6efd; text-decoration: underline; font-weight: bold;">Read our case study &rarr;</a>' +
                    '</div>' +
                '</div><p> </p>'
        },

        // 4. TEXT LEFT, IMAGE RIGHT
        {
            title: "Split: Image Right",
            image: "iR.jpg",
            description: "Descriptive text on the left, large image on the right.",
            html: 
                '<div style="display: flex; flex-wrap: wrap-reverse; align-items: center; gap: 40px; margin-bottom: 40px;">' +
                    '<div style="flex: 1; min-width: 250px;">' +
                        '<h2 style="font-size: 28px; margin-bottom: 15px;">Seamless Integration</h2>' +
                        '<p style="font-size: 16px; line-height: 1.6; color: #555;">Connect your favorite tools in minutes. Our API-first approach ensures you stay flexible and fast.</p>' +
                        '<a href="#" style="color: #0d6efd; text-decoration: underline; font-weight: bold;">See all integrations &rarr;</a>' +
                    '</div>' +
                    '<div style="flex: 1; min-width: 250px;">' +
                        '<img src="https://quickersite.com/placeholder.jpg" style="width:100%; border-radius: 12px;" />' +
                    '</div>' +
                '</div><p> </p>'
        },
		{
			title: "Gradient Call-to-Action",
			image: "gradient.jpg",
			description: "A vibrant gradient banner with a header, tagline, and action link.",
			html: 
				'<div style="background: linear-gradient(135deg, #6610f2 0%, #dc3545 100%); padding: 60px 40px; border-radius: 12px; text-align: center; color: #ffffff;  margin-bottom: 30px; box-shadow: 0 10px 20px rgba(13, 110, 253, 0.2);">' +
					'<h2 style="font-size: 36px; margin: 0 0 15px 0; color: #ffffff;">Ready to Transform Your Workflow?</h2>' +
					'<p style="font-size: 18px; margin: 0 0 30px 0; opacity: 0.9; max-width: 700px; margin-left: auto; margin-right: auto;">' +
						'Join over 10,000 teams using our platform to streamline their creative process and deliver faster results.' +
					'</p>' +
					'<div style="font-size: 20px;">' +
						'<a href="#" style="color: #ffffff; text-decoration: none; font-weight: bold; border-bottom: 2px solid #ffffff; padding-bottom: 4px; transition: opacity 0.2s;">' +
							'Get started today &rarr;' +
						'</a>' +
					'</div>' +
				'</div><p> </p>'
		},
		{
			title: "Testimonial Card",
			image: "testimonial.jpg",
			description: "Single testimonial with quote, photo, name, and title.",
			html: 
				'<div style="background: #f8f9fa; border-left: 4px solid #0d6efd; border-radius: 8px; padding: 30px; margin-bottom: 30px; max-width: 600px;">' +
					'<div style="margin-bottom: 20px;">' +
						'<svg width="40" height="40" viewBox="0 0 40 40" style="fill: #0d6efd; opacity: 0.3;">' +
							'<path d="M10 16c0-3.866 3.134-7 7-7v4c-1.654 0-3 1.346-3 3v1h3v10h-7V16zm16 0c0-3.866 3.134-7 7-7v4c-1.654 0-3 1.346-3 3v1h3v10h-7V16z"/>' +
						'</svg>' +
					'</div>' +
					'<p style="font-size: 18px; line-height: 1.7; color: #333; margin: 0 0 25px 0; font-style: italic;">' +
						'"This product completely transformed how our team works. The intuitive interface and powerful features saved us countless hours every week."' +
					'</p>' +
					'<div style="display: flex; align-items: center; gap: 15px;">' +
						'<img src="https://quickersite.com/placeholder.jpg" alt="Customer photo" style="width: 60px; height: 60px; border-radius: 50%; object-fit: cover;" />' +
						'<div>' +
							'<div style="font-weight: bold; color: #222; font-size: 16px;">Sarah Johnson</div>' +
							'<div style="color: #666; font-size: 14px;">Head of Marketing, TechCorp</div>' +
						'</div>' +
					'</div>' +
				'</div><p> </p>'
		},
		{
			title: "Responsive Video Embed",
			image: "video.jpg",
			description: "16:9 responsive YouTube/Vimeo video player that scales perfectly.",
			html: 
				'<div style="position: relative; padding-bottom: 56.25%; height: 0; overflow: hidden; max-width: 100%; margin-bottom: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">' +
					'<iframe style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; border: 0;" src="https://www.youtube.com/embed/dQw4w9WgXcQ" title="Video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>' +
				'</div><p> </p>'
		},		
		
        // 5. EMAIL-SAFE TEMPLATE (TABLE BASED)
        {
			title: "Email Safe: Hero Header",
			image: "mail.jpg",
			description: "Table-based layout with a background image header (Outlook safe).",
			html: 
				'<div align="center" style="background-color: #f4f4f4; padding: 20px 0;">' +
					'' +
					'<table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; font-family: Arial, sans-serif; border-collapse: collapse;">' +
						'<tr>' +
							'<td align="center" valign="top" style="position: relative;">' +
								'' +
								'' +
								'<div style="background-image: url(\'https://images.unsplash.com/photo-1557683316-973673baf926?auto=format&fit=crop&w=600&h=300&q=80\'); background-size: cover; background-position: center; height: 300px; display: table; width: 100%;">' +
									'<div style="display: table-cell; vertical-align: middle; text-align: center; background: rgba(0,0,0,0.4);">' +
										'<h1 style="color: #ffffff; font-size: 36px; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.5);">Main Title Here</h1>' +
										'<p style="color: #ffffff; font-size: 18px; margin: 10px 0 0 0; opacity: 0.9;">Professional Subtitle or Catchphrase</p>' +
									'</div>' +
								'</div>' +
								'' +
							'</td>' +
						'</tr>' +
						'' +
						'<tr>' +
							'<td style="padding: 40px 30px; line-height: 1.6; color: #444444; font-size: 16px;">' +
								'<h3>Newsletter Headline</h3>' +
								'<p>This template uses a hybrid approach. The background image will render in Gmail/Apple Mail via CSS, while the VML code ensures Outlook users see the image too. You can safely edit the text overlays above.</p>' +
								'<a href="#" style="color: #0d6efd; font-weight: bold; text-decoration: underline;">Read the full story &rarr;</a>' +
							'</td>' +
						'</tr>' +
						'' +
						'<tr>' +
							'<td bgcolor="#333333" style="padding: 20px; color: #999999; font-size: 14px; text-align: center;">' +
								'You are receiving this because you subscribed via our website.<br/>' +
								'<a href="#" style="color: #ffffff; text-decoration: underline;">Unsubscribe</a>' +
							'</td>' +
						'</tr>' +
					'</table>' +
				'</div><p> </p>'
		}
    ]
});