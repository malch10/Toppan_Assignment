use strict;
use warnings;
use XML::LibXML;
use POSIX qw(strftime);
use Date::Parse;

my $orderform_content;
{
    local $/;
    open my $fh, '<', 'orderform.txt' or die "Could not open 'orderform.txt' $!";
    $orderform_content = <$fh>;
    close $fh;
}


my ($customer_name) = $orderform_content =~ /(!CUSTOMER_NAME__!!)/;
my ($date_order) = $orderform_content =~ /DATE\s*:\s*(\d{2}\/\d{2}\/\d{4})/;
my ($site_of) = $orderform_content =~ /SITE OF\s*:\s*(\w+)/;
my ($orderform_number) = $orderform_content =~ /ORDERFORM NUMBER\s*:\s*([\w]+\s*(\d+))/;
my ($revision) = $orderform_content =~ /REVISION\s*:\s*(\d+)/;
my ($page) = $orderform_content =~ /PAGE\s*:\s*(\d+)/;
my ($technology_name) = $orderform_content =~ /TECHNOLOGY NAME\s*:\s*([\w]+\s*[A-Z0-9.%]+)/;
my ($status) = $orderform_content =~ /STATUS\s*:\s*(\w+)/;
my ($mask_set_name) = $orderform_content =~ /MASK SET NAME\s*:\s*(\w+)/;
my ($fab_unit) = $orderform_content =~ /FAB UNIT\s*:\s*(\w+)/;
my ($email_address) = $orderform_content =~ /EMAIL\s*:\s*([\w\@\:\.]+)/;
my ($po_numbers) = $orderform_content =~ /P.O. NUMBERS\s*:\s*(\w+)/;
my ($site_to_send_masks_to) = $orderform_content =~ /SITE TO SEND MASKS TO\s*:\s*(\w+)/;
my ($site_to_send_invoice_to) = $orderform_content =~ /SITE TO SEND INVOICE TO\s*:\s*(\w+\s*\w+)/;
my ($technical_contact) = $orderform_content =~ /TECHNICAL CONTACT\s*:\s*([\w]+\s*[A-Z0-9.%\/]+)/;
my ($device) = $orderform_content =~ /DEVICE\s*:\s*([\w@]+)/;


$date_order = strftime('%Y-%m-%d', localtime(str2time($date_order)));

my $doc = XML::LibXML::Document->new('1.0', 'UTF-8');
my $order_form = $doc->createElement('OrderForm');

$order_form->appendTextChild('Customer', $customer_name);
$order_form->appendTextChild('Device', $device);
$order_form->appendTextChild('MaskSupplier', 'TOPPAN');
$order_form->appendTextChild('Date', $date_order);
$order_form->appendTextChild('SiteOf', $site_of);
$order_form->appendTextChild('OrderFormNumber', $orderform_number);
$order_form->appendTextChild('Revision', $revision);
$order_form->appendTextChild('Page', $page);
$order_form->appendTextChild('TechnologyName', $technology_name);
$order_form->appendTextChild('Status', $status);
$order_form->appendTextChild('MaskSetName', $mask_set_name);
$order_form->appendTextChild('FabUnit', $fab_unit);
$order_form->appendTextChild('EmailAddress', $email_address);

my $levels = $order_form->appendChild($doc->createElement('Levels'));
while ($orderform_content =~ /(\d{2})\s*\|([A-Z\s\d]+?)\s*\|\s*(\d+)\s*\|\s*([\w]+)\s*\|\s*(\d+)\s*\|\s*(\d{2}\w{3}\d{2})/g) {
    my $level = $levels->appendChild($doc->createElement('Level'));
    $level->setAttribute('num', $1);
    $level->appendTextChild('MaskCodification', $2);
    $level->appendTextChild('Group', $3);
    $level->appendTextChild('Cycle', $4);
    $level->appendTextChild('Quantity', $5);
    my $shipdate = strftime('%Y-%m-%d', localtime(str2time($6)));
    $level->appendTextChild('ShipDate', $shipdate);
}

my $cd_information = $order_form->appendChild($doc->createElement('Cdinformation'));
while ($orderform_content =~ /\|\s*(\d{2})\s*\|\s*(\w{2})\s*\|\s*([GF])\s*\|\s*(\w{2})\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*(\d+)\s*\|\s*(\w+)\s*\|\s*(\w+)\s*\|\s*(\w+)\s*\|\s*(\w+)\s*\|\s*(\w+)\s*\|\s*(\w+)\s*\|\s*(\w+)\s*\|\s*([\w.]+)\s*\|\s*(\d+)\s*\|\s*(\w{3}\d{2}-\d{2}:\d{2})\s*\|/g) {
    my $cd_level = $cd_information->appendChild($doc->createElement('Level'));
    $cd_level->appendTextChild('Revision', $2);
    $cd_level->appendTextChild('CDNumber', $7);
    $cd_level->appendTextChild('CDName', $8);
    $cd_level->appendTextChild('Feature', $9);
    $cd_level->appendTextChild('Tone', $10);
    $cd_level->appendTextChild('Polarity', $11);
}

$order_form->appendTextChild('PONumbers', $po_numbers);
$order_form->appendTextChild('SiteToSendMasksTo', $site_to_send_masks_to);
$order_form->appendTextChild('SiteToSendInvoiceTo', $site_to_send_invoice_to);
$order_form->appendTextChild('TechnicalContact', $technical_contact);
$order_form->appendTextChild('ShippingMethod', '');
$order_form->appendTextChild('AdditionalInformation', '');

$doc->setDocumentElement($order_form);
$doc->toFile('output_perl.xml', 1);
