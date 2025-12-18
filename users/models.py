from django.contrib.auth.models import AbstractUser, BaseUserManager
from django.db import models
from django.utils.translation import gettext_lazy as _

class CustomUserManager(BaseUserManager):
    """
    Custom user manager where email is the unique identifier
    instead of usernames.
    """
    def create_user(self, email, password=None, **extra_fields):
        if not email:
            raise ValueError(_('The Email must be set'))
        email = self.normalize_email(email)
        
        # Default new users to STAFF role unless specified
        extra_fields.setdefault('role', User.Role.STAFF)
        
        user = self.model(email=email, **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, email, password=None, **extra_fields):
        extra_fields.setdefault('is_staff', True)
        extra_fields.setdefault('is_superuser', True)
        extra_fields.setdefault('is_active', True)
        extra_fields.setdefault('role', User.Role.ADMIN)

        if extra_fields.get('is_staff') is not True:
            raise ValueError(_('Superuser must have is_staff=True.'))
        if extra_fields.get('is_superuser') is not True:
            raise ValueError(_('Superuser must have is_superuser=True.'))

        return self.create_user(email, password, **extra_fields)

class User(AbstractUser):
    class Role(models.TextChoices):
        ADMIN = 'ADMIN', _('System Admin')
        OFFICER = 'OFFICER', _('Sectional Officer')
        CSI_SUPERVISOR = 'CSI', _('CSI Supervisor')
        STAFF = 'STAFF', _('Field Staff')
        VIEWER = 'VIEWER', _('Viewer/Auditor')

    # Remove the username field
    username = None
    
    # Make email unique and the identifier
    email = models.EmailField(_('email address'), unique=True)

    # Custom fields
    role = models.CharField(max_length=20, choices=Role.choices, default=Role.STAFF)
    
    # Link to Master Data for auto-filling forms
    designation = models.ForeignKey(
        'office.Designation', 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True
    )
    posting_station = models.ForeignKey(
        'office.Station',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='posted_staff'
    )
    phone_number = models.CharField(max_length=15, blank=True)

    # Django Auth Settings
    USERNAME_FIELD = 'email'
    REQUIRED_FIELDS = [] # Email is required by default, no other fields needed for createsuperuser

    objects = CustomUserManager()

    def __str__(self):
        return f"{self.email} ({self.role})"

# from django.db import models
# from django.contrib.auth.models import AbstractUser
# from django.utils.translation import gettext_lazy as _

# class User(AbstractUser):
#     """
#     Custom User model supporting roles and linkage to specific 
#     stations or CSIs for auto-filling forms.
#     """
#     class Role(models.TextChoices):
#         ADMIN = 'ADMIN', _('System Admin')
#         OFFICER = 'OFFICER', _('Sectional Officer')
#         CSI_SUPERVISOR = 'CSI', _('CSI Supervisor')
#         STAFF = 'STAFF', _('Field Staff')
#         VIEWER = 'VIEWER', _('Viewer/Auditor')

#     role = models.CharField(
#         max_length=20, 
#         choices=Role.choices, 
#         default=Role.STAFF,
#         help_text="Controls broad access permissions."
#     )
    
#     # Link to Master Data (Optional, allows auto-filling forms)
#     # Using string reference 'office.Designation' to avoid circular imports
#     designation = models.ForeignKey(
#         'office.Designation', 
#         on_delete=models.SET_NULL, 
#         null=True, 
#         blank=True
#     )
    
#     # Where is this user physically posted?
#     posting_station = models.ForeignKey(
#         'office.Station',
#         on_delete=models.SET_NULL,
#         null=True,
#         blank=True,
#         related_name='posted_staff'
#     )

#     phone_number = models.CharField(max_length=15, blank=True)
    
#     def __str__(self):
#         return f"{self.first_name} {self.last_name} ({self.username}) - {self.role}"