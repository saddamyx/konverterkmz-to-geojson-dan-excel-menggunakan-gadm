import streamlit as st
import zipfile
import os
import geojson
import xml.etree.ElementTree as ET
import pandas as pd
from pyproj import Proj
import geopandas as gpd
from shapely.geometry import Point, Polygon
from io import BytesIO

# Fungsi untuk mendekompres file KMZ dan mengambil file KML
def extract_kml_from_kmz(kmz_file, output_folder):
    try:
        with zipfile.ZipFile(kmz_file, 'r') as kmz:
            kmz.extractall(output_folder)
            for file in os.listdir(output_folder):
                if file.endswith(".kml"):
                    return os.path.join(output_folder, file)
    except Exception as e:
        st.error(f"Error extracting KML: {e}")
    return None

# Fungsi untuk mencocokkan koordinat dengan data GADM
def reverse_geocode(lat, lon, gadm_gdf):
    try:
        point = Point(lon, lat)
        match = gadm_gdf[gadm_gdf.contains(point)]

        if not match.empty:
            desa = match.iloc[0].get("NAME_4", "Tidak diketahui")
            kecamatan = match.iloc[0].get("NAME_3", "Tidak diketahui")
            kabupaten = match.iloc[0].get("NAME_2", "Tidak diketahui")
            provinsi = match.iloc[0].get("NAME_1", "Tidak diketahui")
        else:
            desa, kecamatan, kabupaten, provinsi = "Tidak diketahui", "Tidak diketahui", "Tidak diketahui", "Tidak diketahui"

        return desa, kecamatan, kabupaten, provinsi
    except Exception as e:
        st.warning(f"Reverse geocoding error: {e}")
        return "Tidak diketahui", "Tidak diketahui", "Tidak diketahui", "Tidak diketahui"

# Fungsi untuk konversi koordinat ke UTM
def latlon_to_utm(lat, lon):
    try:
        zone_number = int((lon + 180) // 6) + 1
        proj_utm = Proj(proj="utm", zone=zone_number, ellps="WGS84")
        utm_x, utm_y = proj_utm(lon, lat)
        
        # Memastikan Northing selalu positif
        if utm_y < 0:
            utm_y += 10_000_000  # Menyesuaikan untuk belahan bumi selatan (konvensi UTM)
        
        return round(utm_x, 2), round(utm_y, 2), zone_number
    except Exception as e:
        st.warning(f"UTM conversion error: {e}")
        return None, None, None


# Fungsi untuk menghitung luas poligon dalam hektar
def calculate_area(coords):
    try:
        # Membentuk poligon dari koordinat
        polygon = Polygon(coords)

        # Membuat GeoDataFrame untuk poligon dengan CRS WGS84
        gdf = gpd.GeoDataFrame(index=[0], crs="EPSG:4326", geometry=[polygon])

        # Mengubah CRS ke UTM secara otomatis berdasarkan lokasi
        gdf = gdf.to_crs(gdf.estimate_utm_crs())

        # Menghitung luas dalam meter persegi dan konversi ke hektar
        area_m2 = gdf.geometry.area[0]
        area_ha = area_m2 / 10000  # Konversi dari m² ke hektar
        return round(area_ha, 2)
    except Exception as e:
        st.warning(f"Error calculating area: {e}")
        return None


# Fungsi kml_to_geojson
def kml_to_geojson(kml_file, kmz_filename, gadm_gdf):
    try:
        tree = ET.parse(kml_file)
        root = tree.getroot()
        namespace = {'ns': 'http://www.opengis.net/kml/2.2'}

        geojson_data = {
            "type": "FeatureCollection",
            "features": []
        }

        for placemark in root.findall('.//ns:Placemark', namespace):
            name = placemark.find('./ns:name', namespace).text if placemark.find('./ns:name', namespace) is not None else "Unnamed"
            coordinates = placemark.find('.//ns:coordinates', namespace).text if placemark.find('.//ns:coordinates', namespace) is not None else None

            if coordinates:
                coords = coordinates.strip().split(' ')
                coords = [coord.split(',') for coord in coords]
                coords = [[float(coord[0]), float(coord[1])] for coord in coords]

                if coords:
                    lon, lat = coords[0][0], coords[0][1]
                    easting, northing, utm_zone = latlon_to_utm(lat, lon)
                    desa, kecamatan, kabupaten, provinsi = reverse_geocode(lat, lon, gadm_gdf)
                    area_ha = calculate_area(coords)  # Memperbaiki perhitungan luas

                    feature = {
                        "type": "Feature",
                        "geometry": {
                            "type": "Polygon",
                            "coordinates": [coords]
                        },
                        "properties": {
                            "name": name,
                            "kmz_filename": kmz_filename,
                            "coordinates_dd": coords,
                            "Easting (UTM)": easting,
                            "Northing (UTM)": northing,
                            "UTM Zone": f"{utm_zone}N",
                            "Desa": desa,
                            "Kecamatan": kecamatan,
                            "Kabupaten": kabupaten,
                            "Provinsi": provinsi,
                            "Luas (Ha)": area_ha
                        }
                    }
                    geojson_data["features"].append(feature)

        return [(f"{kmz_filename}.geojson", geojson.dumps(geojson_data))]
    except Exception as e:
        st.error(f"Error converting KML to GeoJSON: {e}")
        return []


# Fungsi untuk mengonversi GeoJSON menjadi tabel Excel
def geojson_to_excel(geojson_files):
    data = []

    for geojson_filename, geojson_data in geojson_files:
        features = geojson.loads(geojson_data).get("features", [])

        for feature in features:
            properties = feature.get('properties', {})
            coords = feature.get('geometry', {}).get('coordinates', [[]])[0]

            for lon, lat in coords:
                easting, northing, utm_zone = latlon_to_utm(lat, lon)
                data.append({
                    "GeoJSON Filename": geojson_filename,
                    "Name": properties.get('name', 'Unnamed'),
                    "Longitude (DD)": lon,
                    "Latitude (DD)": lat,
                    "Easting (UTM)": easting,
                    "Northing (UTM)": northing,
                    "UTM Zone": f"{utm_zone}N",
                    "Desa": properties.get('Desa', 'Unknown'),
                    "Kecamatan": properties.get('Kecamatan', 'Unknown'),
                    "Kabupaten": properties.get('Kabupaten', 'Unknown'),
                    "Provinsi": properties.get('Provinsi', 'Unknown'),
                    "Luas (Ha)": properties.get('Luas (Ha)', 'Unknown')
                })

    return pd.DataFrame(data)

# Fungsi utama aplikasi Streamlit
def main():
    st.title("KMZ to GeoJSON and Excel Converter with GADM")

    # Load GADM GeoJSON
    gadm_path = "gadm41_IDN_4.json"
    if not os.path.exists(gadm_path):
        st.error("GADM data file not found.")
        return

    gadm_gdf = gpd.read_file(gadm_path)

    uploaded_file = st.file_uploader("Upload KMZ file", type=["kmz"])

    if uploaded_file:
        with st.spinner("Processing KMZ file..."):
            temp_folder = "temp_kmz"
            os.makedirs(temp_folder, exist_ok=True)

            kmz_path = os.path.join(temp_folder, "uploaded.kmz")
            with open(kmz_path, "wb") as f:
                f.write(uploaded_file.read())

            kml_file = extract_kml_from_kmz(kmz_path, temp_folder)
            if not kml_file:
                st.error("Failed to extract KML from KMZ file.")
                return

            kmz_filename = uploaded_file.name
            geojson_files = kml_to_geojson(kml_file, kmz_filename, gadm_gdf)

            if geojson_files:
                geojson_df = geojson_to_excel(geojson_files)

                st.write("Location Information:")
                st.dataframe(geojson_df)

                # Create a ZIP file for GeoJSON outputs
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for geojson_filename, geojson_data in geojson_files:
                        zipf.writestr(geojson_filename, geojson_data)

                zip_buffer.seek(0)
                st.download_button(
                    label="Download GeoJSON Files (ZIP)",
                    data=zip_buffer,
                    file_name="geojson_files.zip",
                    mime="application/zip"
                )

                # Create an Excel file for location data
                excel_file = BytesIO()
                with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                    geojson_df.to_excel(writer, index=False, sheet_name='Locations')

                excel_file.seek(0)
                st.download_button(
                    label="Download Excel File",
                    data=excel_file,
                    file_name=f"{kmz_filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.error("No GeoJSON data was generated. Please check your KMZ file.")

            # Clean up temporary folder
            if os.path.exists(temp_folder):
                for file in os.listdir(temp_folder):
                    file_path = os.path.join(temp_folder, file)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                os.rmdir(temp_folder)

if __name__ == "__main__":
    main()


