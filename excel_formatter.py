import pandas as pd
import numpy as np

class CreatePazaramaFormattedExcel:
    missing_columns_dict = {
        'charm': {
            'category_id': 'e114f605-160e-4de8-e8b5-08dc5d2238d7',
            'missing_columns': []
        },
        'bileklik': {
            'category_id': '10745b70-5018-4913-871d-c44669a1de79',
            'missing_columns': ['Renk', 'Materyal', 'Ölçü', 'Kapama Türü', 'Taş Cinsi', 'Pırlanta Rengi']
        },
        'bilezik': {
            'category_id': 'ac19c440-9234-40a0-8835-3625ddcd462f',
            'missing_columns': ['Renk', 'Cinsiyet', 'Materyal', 'Ölçü', 'Taş Cinsi', 'Tip', 'Pırlanta Rengi']
        },
        'kolye': {
            'category_id': '653aaba3-783e-4da5-b9ef-c1b029a97d9a',
            'missing_columns': ['Renk', 'Materyal', 'Taş Cinsi', 'Uzunluk', 'Pırlanta Rengi']
        },
        'küpe': {
            'category_id': 'fd561a50-5469-4a86-a9f7-c88b82f41805',
            'missing_columns': ['Renk', 'Materyal', 'Model', 'Taş Cinsi', 'Pırlanta Rengi']
        },
        'takı seti': {
            'category_id': '5d19116d-1c96-45f1-bb97-08241afa7452',
            'missing_columns': ['Yüzük Ölçüsü', 'Taş Cinsi', 'Taş Rengi', 'Pırlanta Rengi']
        },
        'köpek elbisesi': {
            'category_id': 'ab45c125-1fef-454e-87e6-ba6154273c86',
            'missing_columns': ['Renk', 'Köpek Irkı', 'Beden/Yaş']
        }
    }

    def __init__(self, category:str, filename:str) -> None:
        self.category = category
        self.category_id = self.missing_columns_dict[self.category.lower()]['category_id']
        self.filename = filename
        self.df = pd.read_excel(f'trendyol/{self.filename}')
        self.df['Barkod'] = self.df['Barkod'].astype(str)

        # run program
        self.main()


    def drop_unnecessary_columns(self):
        drop_list = ['Partner ID', 'Komisyon Oranı', 'Cinsiyet', 'Boyut/Ebat', 'Kategori İsmi', "Trendyol'da Satılacak Fiyat (KDV Dahil)", 'BuyBox Fiyatı', 'Desi', 'Sevkiyat Süresi', 'Sevkiyat Tipi', 'Durum', 'Ne Yapmalıyım', 'Trendyol.com Linki']

        for col in drop_list:
            self.df.drop(col, axis=1, inplace=True)

    
    def edit_photo_columns(self):
        photos_df = self.df[['Görsel 1', 'Görsel 2', 'Görsel 3', 'Görsel 4', 'Görsel 5']]
        photos_df.columns = ['Görsel Linki-1', 'Görsel Linki-2', 'Görsel Linki-3', 'Görsel Linki-4', 'Görsel Linki-5']

        # drop ex photo columns
        for col in ['Görsel 1', 'Görsel 2', 'Görsel 3', 'Görsel 4', 'Görsel 5', 'Görsel 6', 'Görsel 7', 'Görsel 8']:
            self.df.drop(col, axis=1, inplace=True)

        # concat new photo columns and df
        self.df = pd.concat([self.df, photos_df], axis=1)


    def rename_columns(self):
        self.df.rename(columns={'Model Kodu':'Grup Kodu',
           'Ürün Rengi':'Renk',
           'Beden':'Ölçü',
           'Tedarikçi Stok Kodu':'Stok Kodu',
           'Ürün Açıklaması':'Ürün Açıklama',
           'Piyasa Satış Fiyatı (KDV Dahil)':'Satış Fiyatı',
           'Ürün Stok Adedi':'Stok Adedi'}, inplace=True)
        

    def append_missing_columns(self):
        missing_columns = {
            'Kategori': [f'{self.category_id}' for x in range(len(self.df))],
            'İndirimli Satış Fiyatı':self.df['Satış Fiyatı'].copy(),
            'Para Birimi': ['TRY' for x in range(len(self.df))],
            'Teslimat Seçeneği':['' for x in range(len(self.df))],
            'Teslim İl':['' for x in range(len(self.df))]
        }

        for col in self.missing_columns_dict[self.category.lower()]['missing_columns']:
            if col not in list(self.df.columns):
                missing_columns[col] = ['' for x in range(len(self.df))]

        missing_columns_df = pd.DataFrame(missing_columns)

        # concat missing_columns_df and df
        self.df = pd.concat([self.df, missing_columns_df], axis=1)


    def sort_df(self):
        desired_order = ['Barkod', 'Marka', 'Grup Kodu', 'Kategori', 'Para Birimi', 'Ürün Adı',
            'Ürün Açıklama', 'Satış Fiyatı', 'İndirimli Satış Fiyatı', 'Stok Adedi', 'Stok Kodu', 'KDV Oranı', 'Görsel Linki-1', 'Görsel Linki-2', 'Görsel Linki-3', 'Görsel Linki-4', 'Görsel Linki-5', 'Teslimat Seçeneği',
            'Teslim İl'] + self.missing_columns_dict[self.category.lower()]['missing_columns']
        
        self.df = self.df[desired_order]


    def fix_brand_names(self):
        if 'Marka' in list(self.df.columns):
            for index in range(len(self.df)):
                brand = str(self.df.loc[index, 'Marka'])
                if brand == 'takantakana':
                    self.df.loc[index, 'Marka'] = brand.capitalize()
                elif brand == 'Worldshop':
                    self.df.loc[index, 'Marka'] == brand.lower()


    def fix_measurements(self):
        if 'Ölçü' in list(self.df.columns):
            self.df['Ölçü'].fillna('Standart', inplace=True)
            for index in range(len(self.df)):
                if str(self.df.loc[index, 'Ölçü']) == 'Tek Ebat':
                    self.df.loc[index, 'Ölçü'] = 'Standart'


    def fix_colors(self):
        if 'Renk' in list(self.df.columns):
            for index in range(len(self.df)):
                color = str(self.df.loc[index, 'Renk'])
                if color == 'gümüş':
                    self.df.loc[index, 'Renk'] = 'Gümüş'
                
                elif color == 'GÜMÜŞ':
                    self.df.loc[index, 'Renk'] = 'Gümüş'

                elif color == 'altın':
                    self.df.loc[index, 'Renk'] = 'Altın'

                elif color == 'altın kaplama':
                    self.df.loc[index, 'Renk'] = 'Altın'

                elif color == 'ALTIN KAPLAMA':
                    self.df.loc[index, 'Renk'] = 'Altın'

                elif color == 'mavi':
                    self.df.loc[index, 'Renk'] = 'Mavi'

                elif color == 'LİLA':
                    self.df.loc[index, 'Renk'] = 'Lila'

                elif color == 'Karışık':
                    self.df.loc[index, 'Renk'] = 'Çok Renkli'

                elif color == 'Renkli':
                    self.df.loc[index, 'Renk'] = 'Çok Renkli'

                elif color == 'Erkek':
                    self.df.loc[index, 'Renk'] = ''

                elif color == 'Kadın':
                    self.df.loc[index, 'Renk'] = ''


    def append_starter_to_prod_name(self):
        if self.category.lower() == 'bileklik':
            for index in range(len(self.df)):
                prod_name = self.df.loc[index, 'Ürün Adı']
                if not prod_name.startswith('Pandora Tarz'):
                    self.df.loc[index, 'Ürün Adı'] = f'Pandora Tarz, {prod_name}'


    def convert_to_excel(self):
        print(f'{self.filename} named file saved under /pazarama directory!')
        self.df.to_excel(f'pazarama/edited_{self.filename}', index=False)


    def main(self):
        self.drop_unnecessary_columns()
        self.edit_photo_columns()
        self.rename_columns()
        self.append_missing_columns()
        self.sort_df()
        self.fix_brand_names()
        self.fix_measurements()
        self.fix_colors()
        self.append_starter_to_prod_name()
        self.convert_to_excel()


if __name__ == '__main__':
    category_list = ['Bileklik', 'Charm', 'Bilezik', 'Kolye', 'Küpe', 'Takı Seti', 'Köpek Elbisesi']
    filename_list = ['Bileklik.xlsx', 'Charm.xlsx', 'Bilezik.xlsx', 'Kolye.xlsx', 'Küpe.xlsx', 'Takı Seti.xlsx', 'Köpek Elbisesi.xlsx']
    for category, filename in zip(category_list, filename_list):
        excel_formatter = CreatePazaramaFormattedExcel(
            category=category,
            filename=filename
        )