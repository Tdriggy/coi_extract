import glob
import os
import pdfquery
import clipboard
import time
import xpath_address_map
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


class AccordReader:

    def __init__(self, file_name):
        self.file_name = file_name
        self.url = clipboard.paste()

    coordinate_dict = {'insurer_a':        (4.82, 8.42, 7.03, 8.52),
                       'insurer_b':        (4.82, 8.27, 7.03, 8.37),
                       'insurer_c':        (4.82, 8.10, 7.03, 8.20),
                       'insurer_d':        (4.82, 7.94, 7.03, 8.04),
                       'insurer_e':        (4.82, 7.77, 7.03, 7.87),
                       'insurer_f':        (4.82, 7.69, 7.03, 7.70),
                       'gl_pol_number':    (2.95, 5.58, 4.58, 6.73),
                       'gl_eff_date':      (4.85, 5.84, 4.98, 6.70),
                       'gl_exp_date':      (5.25, 5.58, 5.80, 6.70),
                       'gl_each_occur':    (7.37, 6.73, 8.17, 6.76),
                       'gl_op_agg':        (7.27, 5.74, 8.24, 5.90),
                       'gl_insurer':       (0.27, 5.91, 0.45, 6.69),
                       'al_pol_number':    (2.95, 4.74, 4.58, 5.57),
                       'al_eff_date':      (4.61, 4.73, 5.25, 5.57),
                       'al_exp_date':      (5.25, 4.73, 5.90, 5.57),
                       'al_comb_limit':    (7.29, 5.40, 8.25, 5.57),
                       'al_insurer':       (0.27, 5.07, 0.43, 5.53),
                       'umb_pol_number':   (2.95, 4.24, 4.61, 4.73),
                       'umb_eff_date':     (4.61, 4.23, 5.25, 4.73),
                       'umb_exp_date':     (5.25, 4.23, 5.90, 4.73),
                       'umb_each_occur':   (7.51, 4.71, 8.15, 4.77),
                       'umb_insurer':      (0.27, 4.41, 0.44, 4.70),
                       'wc_pol_number:':   (3.00, 3.64, 4.52, 4.18),
                       'wc_eff_date':      (4.60, 3.67, 5.25, 4.24),
                       'wc_exp_date':      (5.26, 3.92, 5.90, 4.23),
                       'wc_each_acci':     (7.31, 3.92, 8.25, 4.07),
                       'wc_each_emp':      (7.31, 3.77, 8.24, 3.90),
                       'wc_pol_limit':     (7.43, 3.74, 8.12, 3.76),
                       'wc_insurer':       (0.28, 3.84, 0.43, 4.19),
                       }

    @staticmethod
    def chromedriver_setup():
        my_options = webdriver.ChromeOptions()
        my_options.add_argument('user-data-dir=C:/Users/Tyler/ProgramingProjects/coi_extract/User Data')
        driver_setup = webdriver.Chrome(options=my_options)
        return driver_setup

    # noinspection PyBroadException
    def access_riskonnect(self):
        driver.get(self.url)
        driver.maximize_window()
        WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.ID, 'Login')))
        time.sleep(0.5)
        submit = driver.find_element(By.ID, "Login")
        submit.click()
        try:
            print('Detecting Operation...')
            WebDriverWait(driver, 4).until(ec.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/div/'
                                                                                     'div[2]')))
        except Exception:
            try:
                print('Detecting Operation...')
                WebDriverWait(driver, 3).until(ec.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[1]/'
                                                                                         'section/div[1]/div[2]/div[2]'
                                                                                         '/div[1]/div/div/div/div/div/'
                                                                                         'one-record-home-flexipage2/'
                                                                                         'forcegenerated-adg-rollup_'
                                                                                         'component___force-generated__'
                                                                                         'flexipage_-record-page___-'
                                                                                         'certificate_-record_-page___-'
                                                                                         'certificate_of_-insurance__c'
                                                                                         '___-v-i-e-w/forcegenerated-'
                                                                                         'flexipage_certificate_record_'
                                                                                         'page_certificate_of_insurance'
                                                                                         '__c__view_js/record_flexipage'
                                                                                         '-record-page-decorator/div[1]'
                                                                                         '/records-record-layout-event'
                                                                                         '-broker/slot/slot/flexipage-'
                                                                                         'record-home-template-desktop2'
                                                                                         '/div/div[2]/div[1]/slot/flexi'
                                                                                         'page-component2[1]/slot/flexi'
                                                                                         'page-tabset2/div/lightning-'
                                                                                         'tabset/div/slot/slot/flexi'
                                                                                         'page-tab2[1]/slot/flexipage-'
                                                                                         'component2/slot/records-lwc-'
                                                                                         'detail-panel/records-base-'
                                                                                         'record-form/div/div/div/div/'
                                                                                         'records-lwc-record-layout/'
                                                                                         'forcegenerated-detailpanel_'
                                                                                         'certificate_of_insurance__c'
                                                                                         '___012f1000000n7xuaaq___full'
                                                                                         '___view___recordlayout2/recor'
                                                                                         'ds-record-layout-block/slot/'
                                                                                         'records-record-layout-secti'
                                                                                         'on[1]/div/div/div/slot/recor'
                                                                                         'ds-record-layout-row[1]/slot/'
                                                                                         'records-record-layout-item[1]'
                                                                                         '/div/div/div[1]/span')))
            except Exception:
                print('Detecting Operation...')
                WebDriverWait(driver, 2).until(ec.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[1]/sect'
                                                                                         'ion/div[1]/div[2]/div[2]/div'
                                                                                         '[1]/div/div/div/div/div/one-'
                                                                                         'record-home-flexipage2/force'
                                                                                         'generated-adg-rollup_compone'
                                                                                         'nt___force-generated__flexip'
                                                                                         'age_-record-page___-certific'
                                                                                         'ate_-record_-page___-certifi'
                                                                                         'cate_of_-insurance__c___-v-i-'
                                                                                         'e-w/forcegenerated-flexipage'
                                                                                         '_certificate_record_page_cert'
                                                                                         'ificate_of_insurance__c__vie'
                                                                                         'w_js/record_flexipage-record'
                                                                                         '-page-decorator/div[1]/recor'
                                                                                         'ds-record-layout-event-broke'
                                                                                         'r/slot/slot/flexipage-record'
                                                                                         '-home-template-desktop2/div/'
                                                                                         'div[2]/div[1]/slot/flexipage'
                                                                                         '-component2[1]/slot/flexipag'
                                                                                         'e-tabset2/div/lightning-tabs'
                                                                                         'et/div/slot/slot/flexipage-t'
                                                                                         'ab2[1]/slot/flexipage-compone'
                                                                                         'nt2/slot/records-lwc-detail-p'
                                                                                         'anel/records-base-record-for'
                                                                                         'm/div/div/div/div/records-l'
                                                                                         'wc-record-layout/forcegenera'
                                                                                         'ted-detailpanel_certificate_'
                                                                                         'of_insurance__c___012f1000000'
                                                                                         'n7rpaaa___full___view___recor'
                                                                                         'dlayout2/records-record-layo'
                                                                                         'ut-block/slot/records-record-'
                                                                                         'layout-section[1]/div/div/di'
                                                                                         'v/slot/records-record-layout'
                                                                                         '-row[1]/slot/records-record-l'
                                                                                         'ayout-item[1]/div/div/div[1]'
                                                                                         '/span')))
        time.sleep(1)

    @staticmethod
    def determine_type():
        t1_new_check = driver.find_elements(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div/div/'
                                                      'records-modal-lwc-detail-panel-wrapper/records-record-layout'
                                                      '-event-broker/slot/records-lwc-detail-panel/records-base-record'
                                                      '-form/div/div/div/div/records-lwc-record-layout/forcegenerated'
                                                      '-detailpanel_certificate_of_insurance__c___012f1000000n7xuaaq___'
                                                      'full___create___recordlayout2/records-record-layout-block/slot/'
                                                      'records-record-layout-section[1]/div/div/div/slot/records-record'
                                                      '-layout-row[6]/slot/records-record-layout-item[1]/div/span/slot/'
                                                      'records-record-layout-text-area/lightning-textarea/div/textarea')

        t2_new_check = driver.find_elements(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div/div/'
                                                      'records-modal-lwc-detail-panel-wrapper/records-record-layout'
                                                      '-event-broker/slot/records-lwc-detail-panel/records-base-record'
                                                      '-form/div/div/div/div/records-lwc-record-layout/forcegenerated'
                                                      '-detailpanel_certificate_of_insurance__c___012f1000000n7rpaaa'
                                                      '___full___create___recordlayout2/records-record-layout-block/'
                                                      'slot/records-record-layout-section[1]/div/div/div/slot/records-'
                                                      'record-layout-row[6]/slot/records-record-layout-item[1]/div/span'
                                                      '/slot/records-record-layout-text-area/lightning-textarea/div/'
                                                      'textarea')

        t1_renew_check = driver.find_elements(By.XPATH, '/html/body/div[4]/div[1]/section/div[1]/div[2]/div[2]/div[1]/'
                                                        'div/div/div/div/div/one-record-home-flexipage2/forcegenerated'
                                                        '-adg-rollup_component___force-generated__flexipage_-record-pa'
                                                        'ge___-certificate_-record_-page___-certificate_of_-insurance_'
                                                        '_c___-v-i-e-w/forcegenerated-flexipage_certificate_record_pag'
                                                        'e_certificate_of_insurance__c__view_js/record_flexipage-recor'
                                                        'd-page-decorator/div[1]/records-record-layout-event-broker/sl'
                                                        'ot/slot/flexipage-record-home-template-desktop2/div/div[2]/di'
                                                        'v[1]/slot/flexipage-component2[1]/slot/flexipage-tabset2/div/'
                                                        'lightning-tabset/div/slot/slot/flexipage-tab2[1]/slot/flexipa'
                                                        'ge-component2/slot/records-lwc-detail-panel/records-base-recor'
                                                        'd-form/div/div/div/div/records-lwc-record-layout/forcegenerate'
                                                        'd-detailpanel_certificate_of_insurance__c___012f1000000n7xuaa'
                                                        'q___full___view___recordlayout2/records-record-layout-block/s'
                                                        'lot/records-record-layout-section[1]/div/div/div/slot/records'
                                                        '-record-layout-row[6]/slot/records-record-layout-item[1]/div/d'
                                                        'iv/div[2]/span/slot[1]/records-record-type/div/div')

        t2_renew_check = driver.find_elements(By.XPATH, '/html/body/div[4]/div[1]/section/div[1]/div[2]/div[2]/div[1]/'
                                                        'div/div/div/div/div/one-record-home-flexipage2/forcegenerated'
                                                        '-adg-rollup_component___force-generated__flexipage_-record-pa'
                                                        'ge___-certificate_-record_-page___-certificate_of_-insurance_'
                                                        '_c___-v-i-e-w/forcegenerated-flexipage_certificate_record_pag'
                                                        'e_certificate_of_insurance__c__view_js/record_flexipage-recor'
                                                        'd-page-decorator/div[1]/records-record-layout-event-broker/sl'
                                                        'ot/slot/flexipage-record-home-template-desktop2/div/div[2]/di'
                                                        'v[1]/slot/flexipage-component2[1]/slot/flexipage-tabset2/div/'
                                                        'lightning-tabset/div/slot/slot/flexipage-tab2[1]/slot/flexipa'
                                                        'ge-component2/slot/records-lwc-detail-panel/records-base-reco'
                                                        'rd-form/div/div/div/div/records-lwc-record-layout/forcegenera'
                                                        'ted-detailpanel_certificate_of_insurance__c___012f1000000n7rpa'
                                                        'aa___full___view___recordlayout2/records-record-layout-block/'
                                                        'slot/records-record-layout-section[1]/div/div/div/slot/record'
                                                        's-record-layout-row[1]/slot/records-record-layout-item[1]/div'
                                                        '/div/div[1]/span')

        if t1_new_check:
            print('Operation Detected: New Certificate Tier 1\n')
            return 1
        elif t2_new_check:
            print('Operation Detected: New Certificate Tier 2\n')
            return 2
        elif t1_renew_check:
            print('Operation Detected: Renewed Certificate Tier 1\n')
            return 3
        elif t2_renew_check:
            print('Operation Detected: Renewed Certificate Tier 2\n')
            return 4
        else:
            raise Exception('Unable to detect operation, please try again\n')

    @staticmethod
    def create_map(operation_type_arg):
        map_object = xpath_address_map.XpathMap()
        if operation_type_arg == 1:
            xpath_map_dict = map_object.map_tier1_new
            return xpath_map_dict
        elif operation_type_arg == 2:
            xpath_map_dict = map_object.map_tier2_new
            return xpath_map_dict
        elif operation_type_arg == 3:
            xpath_map_dict = map_object.map_tier1_renew
            return xpath_map_dict
        elif operation_type_arg == 4:
            xpath_map_dict = map_object.map_tier2_renew
            return xpath_map_dict
        else:
            raise Exception('Unable to determine type')

    def pdf_extract_setup(self):
        global pdf
        pdf = pdfquery.PDFQuery(self.file_name)
        pdf.load(0)

    @staticmethod
    def pdf_extract(x1, y1, x2, y2):
        coordinates = (x1 * 72, y1 * 72, x2 * 72, y2 * 72)
        extracted_text = pdf.pq('LTTextLineHorizontal:overlaps_bbox("%d, %d, %d, %d")' % coordinates).text()
        return extracted_text

    def extract_text(self, xpath_map_arg):
        insurer_dict = self.get_insurers()
        self.get_gl_policy_number(self.coordinate_dict['gl_pol_number'], xpath_map_arg)
        self.get_gl_eff_date(self.coordinate_dict['gl_eff_date'], xpath_map_arg)
        self.get_gl_exp_date(self.coordinate_dict['gl_exp_date'], xpath_map_arg)
        self.get_gl_each_occurrence(self.coordinate_dict['gl_each_occur'], xpath_map_arg)
        self.get_gl_op_agg(self.coordinate_dict['gl_op_agg'], xpath_map_arg)
        self.get_gl_insurer(self.coordinate_dict['gl_insurer'], insurer_dict, xpath_map_arg)
        self.get_al_policy_number(self.coordinate_dict['al_pol_number'], xpath_map_arg)
        self.get_al_eff_date(self.coordinate_dict['al_eff_date'], xpath_map_arg)
        self.get_al_exp_date(self.coordinate_dict['al_exp_date'], xpath_map_arg)
        self.get_al_comb_limit(self.coordinate_dict['al_comb_limit'], xpath_map_arg)
        self.get_al_insurer(self.coordinate_dict['al_insurer'], insurer_dict, xpath_map_arg)
        self.get_umb_pol_number(self.coordinate_dict['umb_pol_number'], xpath_map_arg)
        self.get_umb_eff_date(self.coordinate_dict['umb_eff_date'], xpath_map_arg)
        self.get_umb_exp_date(self.coordinate_dict['umb_exp_date'], xpath_map_arg)
        self.get_umb_each_occur(self.coordinate_dict['umb_each_occur'], xpath_map_arg)
        self.get_umb_insurer(self.coordinate_dict['umb_insurer'], insurer_dict, xpath_map_arg)
        self.get_wc_pol_number(self.coordinate_dict['wc_pol_number:'], xpath_map_arg)
        self.get_wc_eff_date(self.coordinate_dict['wc_eff_date'], xpath_map_arg)
        self.get_wc_exp_date(self.coordinate_dict['wc_exp_date'], xpath_map_arg)
        self.get_wc_each_acci(self.coordinate_dict['wc_each_acci'], xpath_map_arg)
        self.get_wc_each_emp(self.coordinate_dict['wc_each_emp'], xpath_map_arg)
        self.get_wc_pol_limit(self.coordinate_dict['wc_pol_limit'], xpath_map_arg)
        self.get_wc_insurer(self.coordinate_dict['wc_insurer'], insurer_dict, xpath_map_arg)

    def get_insurers(self):
        insurer_a = self.pdf_extract(self.coordinate_dict['insurer_a'][0], self.coordinate_dict['insurer_a'][1],
                                     self.coordinate_dict['insurer_a'][2], self.coordinate_dict['insurer_a'][3])
        print('Insurer A: ' + insurer_a)
        insurer_b = self.pdf_extract(self.coordinate_dict['insurer_b'][0], self.coordinate_dict['insurer_b'][1],
                                     self.coordinate_dict['insurer_b'][2], self.coordinate_dict['insurer_b'][3])
        print('Insurer B: ' + insurer_b)
        insurer_c = self.pdf_extract(self.coordinate_dict['insurer_c'][0], self.coordinate_dict['insurer_c'][1],
                                     self.coordinate_dict['insurer_c'][2], self.coordinate_dict['insurer_c'][3])
        print('Insurer C: ' + insurer_c)
        insurer_d = self.pdf_extract(self.coordinate_dict['insurer_d'][0], self.coordinate_dict['insurer_d'][1],
                                     self.coordinate_dict['insurer_d'][2], self.coordinate_dict['insurer_d'][3])
        print('Insurer D: ' + insurer_d)
        insurer_e = self.pdf_extract(self.coordinate_dict['insurer_e'][0], self.coordinate_dict['insurer_e'][1],
                                     self.coordinate_dict['insurer_e'][2], self.coordinate_dict['insurer_e'][3])
        print('Insurer E: ' + insurer_e)
        insurer_f = self.pdf_extract(self.coordinate_dict['insurer_f'][0], self.coordinate_dict['insurer_f'][1],
                                     self.coordinate_dict['insurer_f'][2], self.coordinate_dict['insurer_f'][3])
        print('Insurer F: ' + insurer_f)
        print()
        insurer_dict = {'A': insurer_a, 'B': insurer_b, 'C': insurer_c, 'D': insurer_d, 'E': insurer_e, 'F': insurer_f}
        return insurer_dict

    def get_gl_policy_number(self, coordinates, xpath_map_arg):
        click_button_element_test = driver.find_elements(By.XPATH, xpath_map_arg['click_box'])
        if click_button_element_test:
            click_button_element = driver.find_element(By.XPATH, xpath_map_arg['click_box'])
            click_button_element.click()
        gl_number = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        if gl_number[0:4] in ['X X ', 'x x ']:
            gl_number = gl_number[4:]
        print('GL Number: ' + gl_number)
        try:
            locator = driver.find_element(By.XPATH, xpath_map_arg['gl_locator'])
        except:
            locator = driver.find_element(By.XPATH, xpath_map_arg['gl_locator_2'])
        actions = ActionChains(driver)
        actions.move_to_element(locator).perform()
        gl_pol_number_element = driver.find_element(By.XPATH, xpath_map_arg['gl_pol_number'])
        gl_pol_number_element.clear()
        gl_pol_number_element.send_keys(gl_number)

    def get_gl_eff_date(self, coordinates, xpath_map_arg):
        gl_eff_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('GL Eff Date: ' + gl_eff_date)
        gl_eff_date_element = driver.find_element(By.XPATH, xpath_map_arg['gl_eff_date'])
        gl_eff_date_element.clear()
        gl_eff_date_element.send_keys(gl_eff_date)

    def get_gl_exp_date(self, coordinates, xpath_map_arg):
        gl_exp_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('GL Exp Date: ' + gl_exp_date)
        gl_exp_date_element = driver.find_element(By.XPATH, xpath_map_arg['gl_exp_date'])
        gl_exp_date_element.clear()
        gl_exp_date_element.send_keys(gl_exp_date)

    def get_gl_each_occurrence(self, coordinates, xpath_map_arg):
        gl_each_occurrence = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('GL EO Amount: ' + gl_each_occurrence)
        gl_each_occurrence_element = driver.find_element(By.XPATH, xpath_map_arg['gl_each_occur'])
        gl_each_occurrence_element.clear()
        gl_each_occurrence_element.send_keys(gl_each_occurrence)

    def get_gl_op_agg(self, coordinates, xpath_map_arg):
        gl_op_agg = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('GL Op Agg: ' + gl_op_agg)
        gl_op_agg_element = driver.find_element(By.XPATH, xpath_map_arg['gl_op_agg'])
        gl_op_agg_element.clear()
        gl_op_agg_element.send_keys(gl_op_agg)

    def get_gl_insurer(self, coordinates, insurer_dict, xpath_map_arg):
        gl_insurer = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        if gl_insurer in ['A', 'B', 'C', 'D', 'E', 'F']:
            print('GL Insurer: ' + insurer_dict[gl_insurer])
            gl_insurer_element = driver.find_element(By.XPATH, xpath_map_arg['gl_insurer'])
            gl_insurer_element.clear()
            gl_insurer_element.send_keys(insurer_dict[gl_insurer])
        else:
            pass

    def get_al_policy_number(self, coordinates, xpath_map_arg):
        al_pol_number = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        if al_pol_number[0:4] in ['X X ', 'x x ']:
            al_pol_number = al_pol_number[4:]
        print('AL Number: ' + al_pol_number)
        try:
            locator = driver.find_element(By.XPATH, xpath_map_arg['al_locator'])
        except:
            locator = driver.find_element(By.XPATH, xpath_map_arg['al_locator_2'])
        actions = ActionChains(driver)
        actions.move_to_element(locator).perform()
        al_pol_number_element = driver.find_element(By.XPATH, xpath_map_arg['al_pol_number'])
        al_pol_number_element.clear()
        al_pol_number_element.send_keys(al_pol_number)

    def get_al_eff_date(self, coordinates, xpath_map_arg):
        al_eff_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('AL Eff Date: ' + al_eff_date)
        al_eff_date_element = driver.find_element(By.XPATH, xpath_map_arg['al_eff_date'])
        al_eff_date_element.clear()
        al_eff_date_element.send_keys(al_eff_date)

    def get_al_exp_date(self, coordinates, xpath_map_arg):
        al_exp_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('AL Exp Date: ' + al_exp_date)
        al_exp_date_element = driver.find_element(By.XPATH, xpath_map_arg['al_exp_date'])
        al_exp_date_element.clear()
        al_exp_date_element.send_keys(al_exp_date)

    def get_al_comb_limit(self, coordinates, xpath_map_arg):
        al_comb_limit = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('AL Limit Amount: ' + al_comb_limit)
        al_comb_limit_element = driver.find_element(By.XPATH, xpath_map_arg['al_comb_limit'])
        al_comb_limit_element.clear()
        al_comb_limit_element.send_keys(al_comb_limit)

    def get_al_insurer(self, coordinates, insurer_dict, xpath_map_arg):
        al_insurer = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        if al_insurer in ['A', 'B', 'C', 'D', 'E', 'F']:
            print('AL Insurer: ' + insurer_dict[al_insurer])
            al_insurer_element = driver.find_element(By.XPATH, xpath_map_arg['al_insurer'])
            al_insurer_element.clear()
            al_insurer_element.send_keys(insurer_dict[al_insurer])
        else:
            pass

    def get_umb_pol_number(self, coordinates, xpath_map_arg):
        umb_pol_number = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('UMB Policy Number: ' + umb_pol_number)
        try:
            locator = driver.find_element(By.XPATH, xpath_map_arg['umb_locator'])
        except:
            locator = driver.find_element(By.XPATH, xpath_map_arg['umb_locator_2'])
        actions = ActionChains(driver)
        actions.move_to_element(locator).perform()
        umb_pol_number_element = driver.find_element(By.XPATH, xpath_map_arg['umb_pol_number'])
        umb_pol_number_element.clear()
        umb_pol_number_element.send_keys(umb_pol_number)

    def get_umb_eff_date(self, coordinates, xpath_map_arg):
        umb_eff_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('UMB Eff Date: ' + umb_eff_date)
        umb_eff_date_element = driver.find_element(By.XPATH, xpath_map_arg['umb_eff_date'])
        umb_eff_date_element.clear()
        umb_eff_date_element.send_keys(umb_eff_date)

    def get_umb_exp_date(self, coordinates, xpath_map_arg):
        umb_exp_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('UMB Exp Date: ' + umb_exp_date)
        umb_exp_date_element = driver.find_element(By.XPATH, xpath_map_arg['umb_exp_date'])
        umb_exp_date_element.clear()
        umb_exp_date_element.send_keys(umb_exp_date)

    def get_umb_each_occur(self, coordinates, xpath_map_arg):
        umb_each_occur = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('UMB Each Occur: ' + umb_each_occur)
        umb_each_occur_element = driver.find_element(By.XPATH, xpath_map_arg['umb_each_occur'])
        umb_each_occur_element.clear()
        umb_each_occur_element.send_keys(umb_each_occur)

    def get_umb_insurer(self, coordinates, insurer_dict, xpath_map_arg):
        umb_insurer = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        if umb_insurer in ['A', 'B', 'C', 'D', 'E', 'F']:
            print('UMB Insurer: ' + insurer_dict[umb_insurer])
            umb_insurer_element = driver.find_element(By.XPATH, xpath_map_arg['umb_insurer'])
            umb_insurer_element.clear()
            umb_insurer_element.send_keys(insurer_dict[umb_insurer]) 
        else:
            pass 

    def get_wc_pol_number(self, coordinates, xpath_map_arg):
        wc_pol_number = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        if wc_pol_number != '' and wc_pol_number[0] == 'X':
            wc_pol_number = wc_pol_number[2:]
        print('WC Policy Number: ' + wc_pol_number)
        try:
            locator = driver.find_element(By.XPATH, xpath_map_arg['wc_locator'])
        except:
            locator = driver.find_element(By.XPATH, xpath_map_arg['wc_locator_2'])
        actions = ActionChains(driver)
        actions.move_to_element(locator).perform()
        wc_pol_number_element = driver.find_element(By.XPATH, xpath_map_arg['wc_pol_number'])
        wc_pol_number_element.clear()
        wc_pol_number_element.send_keys(wc_pol_number)

    def get_wc_eff_date(self, coordinates, xpath_map_arg):
        wc_eff_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('WC Eff Date: ' + wc_eff_date)
        wc_eff_date_element = driver.find_element(By.XPATH, xpath_map_arg['wc_eff_date'])
        wc_eff_date_element.clear()
        wc_eff_date_element.send_keys(wc_eff_date)

    def get_wc_exp_date(self, coordinates, xpath_map_arg):
        wc_exp_date = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('WC Exp Date: ' + wc_exp_date)
        wc_exp_date_element = driver.find_element(By.XPATH, xpath_map_arg['wc_exp_date'])
        wc_exp_date_element.clear()
        wc_exp_date_element.send_keys(wc_exp_date)

    def get_wc_each_acci(self, coordinates, xpath_map_arg):
        wc_each_acci = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('wc_each_acci: ' + wc_each_acci)
        wc_each_acci_element = driver.find_element(By.XPATH, xpath_map_arg['wc_each_acci'])
        wc_each_acci_element.clear()
        wc_each_acci_element.send_keys(wc_each_acci)

    def get_wc_each_emp(self, coordinates, xpath_map_arg):
        wc_each_emp = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('WC Each Employee: ' + wc_each_emp)
        wc_each_emp_element = driver.find_element(By.XPATH, xpath_map_arg['wc_each_emp'])
        wc_each_emp_element.clear()
        wc_each_emp_element.send_keys(wc_each_emp)

    def get_wc_pol_limit(self, coordinates, xpath_map_arg):
        wc_pol_limit = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        print('WC Policy Limit: ' + wc_pol_limit)
        wc_pol_limit_element = driver.find_element(By.XPATH, xpath_map_arg['wc_pol_limit'])
        wc_pol_limit_element.clear()
        wc_pol_limit_element.send_keys(wc_pol_limit)

    def get_wc_insurer(self, coordinates, insurer_dict, xpath_map_arg):
        wc_insurer = self.pdf_extract(coordinates[0], coordinates[1], coordinates[2], coordinates[3])
        if wc_insurer in ['A', 'B', 'C', 'D', 'E', 'F']:
            print('WC Insurer: ' + insurer_dict[wc_insurer])
            wc_insurer_element = driver.find_element(By.XPATH, xpath_map_arg['wc_insurer'])
            wc_insurer_element.clear()
            wc_insurer_element.send_keys(insurer_dict[wc_insurer])
        else:
            pass 


if __name__ == '__main__':

    print()
    files_in_downloads = glob. glob('C:/Users/Tyler/Downloads/*.pdf')
    most_recent_file = max(files_in_downloads, key=os. path. getctime)
    reader = AccordReader(most_recent_file)
    driver = reader.chromedriver_setup()
    reader.access_riskonnect()
    operation_type = reader.determine_type()
    xpath_map = reader.create_map(operation_type)
    reader.pdf_extract_setup()
    reader.extract_text(xpath_map)
    print()
