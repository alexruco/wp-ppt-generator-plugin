<?php
/*
Plugin Name: WP PPT Generator Plugin
Description: Generate PowerPoint files from JSON using PHPPresentation.
Version: 1.0
Author: Alex Ruco 
*/

if (!defined('ABSPATH')) exit;

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;

add_action('admin_menu', function() {
    add_management_page(
        'PPT Generator',
        'PPT Generator',
        'manage_options',
        'ppt-generator',
        'ppt_generator_page'
    );
});

function ppt_generator_page() {

    if (isset($_POST['ppt_json'])) {
        $json = stripslashes($_POST['ppt_json']);
        $data = json_decode($json, true);

        if (!$data || !isset($data['slides'])) {
            echo '<div class="notice notice-error"><p>Invalid JSON structure.</p></div>';
        } else {
            $file = ppt_generator_create_ppt($data);
            echo '<div class="notice notice-success"><p>Your PPT is ready: <a href="' . $file . '" download>Download</a></p></div>';
        }
    }

    ?>
    <div class="wrap">
        <h1>PPT Generator from JSON</h1>
        <form method="post">
            <textarea name="ppt_json" rows="12" style="width:100%;font-family:monospace;"><?php
            echo isset($_POST['ppt_json']) ? esc_textarea($_POST['ppt_json']) : '';
            ?></textarea>

            <p><button class="button button-primary">Generate PPT</button></p>
        </form>
    </div>
    <?php
}

function ppt_generator_create_ppt($data) {
    $presentation = new PhpPresentation();
    $presentation->removeSlideByIndex(0);

    foreach ($data['slides'] as $slideData) {
        $slide = $presentation->createSlide();

        if (isset($slideData['title'])) {
            $titleShape = $slide->createRichTextShape();
            $titleShape->setHeight(100);
            $titleShape->setWidth(900);
            $titleShape->setOffsetX(50);
            $titleShape->setOffsetY(40);
            $titleShape->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

            $titleText = $titleShape->createTextRun($slideData['title']);
            $titleText->getFont()->setBold(true);
            $titleText->getFont()->setSize(36);
            $titleText->getFont()->setColor(new Color(Color::COLOR_BLACK));
        }

        if (isset($slideData['text'])) {
            $textShape = $slide->createRichTextShape();
            $textShape->setHeight(300);
            $textShape->setWidth(900);
            $textShape->setOffsetX(50);
            $textShape->setOffsetY(150);

            $bodyText = $textShape->createTextRun($slideData['text']);
            $bodyText->getFont()->setSize(24);
            $bodyText->getFont()->setColor(new Color(Color::COLOR_DARKGRAY));
        }

        if (isset($slideData['bullets']) && is_array($slideData['bullets'])) {
            $bulletShape = $slide->createRichTextShape();
            $bulletShape->setHeight(300);
            $bulletShape->setWidth(900);
            $bulletShape->setOffsetX(80);
            $bulletShape->setOffsetY(260);

            foreach ($slideData['bullets'] as $bullet) {
                $bulletRun = $bulletShape->createTextRun($bullet . "\n");
                $bulletRun->getFont()->setSize(22);
                $bulletRun->getFont()->setColor(new Color(Color::COLOR_BLACK));
                $bulletRun->setBulletStyle(true);
            }
        }

        if (isset($slideData['image'])) {
            $imagePath = __DIR__ . '/' . $slideData['image'];
            if (file_exists($imagePath)) {
                $image = new \PhpOffice\PhpPresentation\Shape\Drawing\File();
                $image->setName("Image");
                $image->setDescription("Image");
                $image->setPath($imagePath);
                $image->setHeight(300);
                $image->setOffsetX(300);
                $image->setOffsetY(200);
                $slide->addShape($image);
            }
        }
    }

    $uploadDir = wp_upload_dir();
    $outputPath = $uploadDir['basedir'] . '/generated-ppt-from-json.pptx';
    $publicUrl = $uploadDir['baseurl'] . '/generated-ppt-from-json.pptx';

    $writer = IOFactory::createWriter($presentation, 'PowerPoint2007');
    $writer->save($outputPath);

    return $publicUrl;
}
