# find_bucket_region.py
import os, boto3
from dotenv import load_dotenv
load_dotenv()

s3 = boto3.client(
    's3',
    aws_access_key_id=os.environ.get('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.environ.get('AWS_SECRET_ACCESS_KEY'),
    region_name='us-east-1',  # use us-east-1 to query any bucket's region
    verify=False
)
bucket = os.environ.get('S3_BUCKET_NAME')
response = s3.get_bucket_location(Bucket=bucket)
print("✅ Actual bucket region:", response['LocationConstraint'])