import boto3
import os
from botocore.exceptions import ClientError
from dotenv import load_dotenv
import io
import certifi
import ssl

load_dotenv()

s3_client = boto3.client(
    's3',
    aws_access_key_id     = os.environ.get('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key = os.environ.get('AWS_SECRET_ACCESS_KEY'),
    region_name           = 'ap-south-2',
    endpoint_url          = 'https://s3.ap-south-2.amazonaws.com',  # ← explicit endpoint
    verify                = False
)
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BUCKET      = os.environ.get('S3_BUCKET_NAME', 'heritage-survey-prod')
S3_PREFIX   = os.environ.get('S3_PREFIX', 'surveys')

# Maps DB column → fixed filename in S3 (matches your screenshot)
PHOTO_FILENAME_MAP = {
    'photo':               'outlet_facade.jpg',
    'interior_photo':      'interior.jpg',
    'shelf_photo':         'shelf_heritage.jpg',
    'posm_photo':          'posm_branding.jpg',
    'cooler_photo_visi':   'cooler_visi.jpg',
    'cooler_photo_bottle': 'cooler_bottle.jpg',
    'cooler_photo_freezer':'cooler_freezer.jpg',
    'space_photo':         'space_available.jpg',
}

def get_s3_key(survey_id, photo_type):
    """
    Returns S3 key like: surveys/TGHSO11000001/outlet_facade.jpg
    """
    filename = PHOTO_FILENAME_MAP.get(photo_type, f'{photo_type}.jpg')
    return f"{S3_PREFIX}/{survey_id}/{filename}"



def upload_to_s3(file_obj, survey_id, photo_type):
    key = get_s3_key(survey_id, photo_type)
    
    # Read bytes and wrap — fixes consumed stream issue
    if hasattr(file_obj, 'read'):
        data = file_obj.read()
    else:
        data = file_obj
    
    s3_client.upload_fileobj(
        io.BytesIO(data),
        BUCKET,
        key,
        ExtraArgs={
            'ContentType': 'image/jpeg',
            'CacheControl': 'max-age=31536000',
        }
    )
    return key

def get_s3_url(s3_key):
    if not s3_key:
        return None

    # Old local filename has no slash — skip it
    if '/' not in s3_key:
        return None

    try:
        url = s3_client.generate_presigned_url(
            'get_object',
            Params={'Bucket': BUCKET, 'Key': s3_key},
            ExpiresIn=604800,
            HttpMethod='GET'
        )
        return url
    except ClientError as e:
        print(f"Presign error for {s3_key}: {e}")  # ← shows error in terminal
        return None


def delete_survey_photos(survey_id):
    """
    Deletes all photos for a survey from S3.
    Called when a survey is deleted.
    """
    prefix = f"{S3_PREFIX}/{survey_id}/"
    try:
        objects = s3_client.list_objects_v2(Bucket=BUCKET, Prefix=prefix)
        if 'Contents' in objects:
            delete_keys = [{'Key': obj['Key']} for obj in objects['Contents']]
            s3_client.delete_objects(Bucket=BUCKET, Delete={'Objects': delete_keys})
    except ClientError as e:
        print(f"S3 delete error for {survey_id}: {e}")